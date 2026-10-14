"""Core scheduling logic for the LOE app (no Streamlit imports, so it is testable on its own).

For every person x budget x milestone the allotted LOE is spread over the months in which
that milestone is active, in whole multiples of a billing increment.  People are independent
so each one is solved separately with OR-tools CP-SAT.
"""
from __future__ import annotations

import datetime as dt
import io
import math
from dataclasses import dataclass, field
from typing import Callable, Iterable

import pandas as pd
from ortools.sat.python import cp_model

OBJECTIVES = {
    "level": "Level each person's load, then spread each milestone evenly",
    "spread": "Spread each milestone evenly only (ignore person-level peaks)",
    "front": "Bill as early as possible",
}


@dataclass
class Rules:
    step: float = 0.5            # billing increment: every monthly entry is a multiple of this
    min_bill: float = 0.5        # smallest non-zero monthly entry (multiple of step)
    cap: float = 21.7            # days / month a person can bill (flag, or hard limit if hard_cap)
    hard_cap: bool = False
    exclude_september: bool = False
    objective: str = "level"
    seconds: float = 5.0         # solver time per stage per person
    workers: int = 8

    def validate(self) -> list[str]:
        msgs = []
        if self.step <= 0: msgs.append("Billing increment must be > 0.")
        if self.min_bill < self.step - 1e-9: msgs.append("Minimum monthly entry cannot be below the billing increment.")
        if abs(self.min_bill / self.step - round(self.min_bill / self.step)) > 1e-9:
            msgs.append("Minimum monthly entry must be a multiple of the billing increment.")
        return msgs


def month_start(d) -> dt.date | None:
    if d is None or (isinstance(d, float) and math.isnan(d)) or pd.isna(d):
        return None
    d = pd.Timestamp(d).date()
    return dt.date(d.year, d.month, 1)


def months_between(start, end) -> list[dt.date]:
    s, e = month_start(start), month_start(end)
    if s is None or e is None or e < s: return []
    out, y, m = [], s.year, s.month
    while (y, m) <= (e.year, e.month):
        out.append(dt.date(y, m, 1)); m += 1
        if m == 13: y, m = y + 1, 1
    return out


def active_map(ms: pd.DataFrame, excluded_months: Iterable[int] = ()) -> dict[tuple[str, str], list[dt.date]]:
    excluded_months = set(excluded_months)
    out = {}
    for r in ms.itertuples():
        if pd.isna(r.budget) or pd.isna(r.milestone): continue
        out[(str(r.budget).strip(), str(r.milestone).strip())] = [
            month for month in months_between(r.start, r.end) if month.month not in excluded_months
        ]
    return out


def clean_loe(loe: pd.DataFrame) -> pd.DataFrame:
    d = loe.copy()
    for c in ("person", "budget", "milestone"): d[c] = d[c].astype("string").str.strip()
    d["loe_days"] = pd.to_numeric(d["loe_days"], errors="coerce").fillna(0.0)
    d = d.dropna(subset=["person", "budget", "milestone"])
    d = d[(d.person != "") & (d.budget != "") & (d.milestone != "") & (d.loe_days > 0)]
    return d.groupby(["person", "budget", "milestone"], as_index=False)["loe_days"].sum()


def input_checks(loe: pd.DataFrame, ms: pd.DataFrame) -> pd.DataFrame:
    """Problems to fix before (or after) optimizing."""
    act = active_map(ms); issues = []
    for r in ms.itertuples():
        if pd.isna(r.budget) or pd.isna(r.milestone): continue
        k = (str(r.budget), str(r.milestone))
        if pd.isna(r.start) or pd.isna(r.end):
            issues.append((k[0], k[1], "Missing dates", "Start and/or end month is blank"))
        elif month_start(r.end) < month_start(r.start):
            issues.append((k[0], k[1], "End before start", f"{r.start} > {r.end}"))
    d = clean_loe(loe)
    for r in d.itertuples():
        if (r.budget, r.milestone) not in act:
            issues.append((r.budget, r.milestone, "LOE on unknown milestone", f"{r.person}: {r.loe_days:g} d but the milestone is not in the Milestones table"))
    have = set(zip(d.budget, d.milestone))
    for k, m in act.items():
        if m and k not in have:
            issues.append((k[0], k[1], "Milestone has no LOE", "Dates are set but nobody has days on it"))
    return pd.DataFrame(issues, columns=["budget", "milestone", "issue", "detail"]).drop_duplicates()


def round_units(cells: dict, step: float, min_units: int) -> dict:
    """Largest-remainder rounding to `step`, keeping the person's total (nearest step).
    Cells that land between 0 and the minimum monthly entry snap to 0 or to the minimum."""
    raw = {k: v / step for k, v in cells.items()}
    units = {k: math.floor(v + 1e-9) for k, v in raw.items()}
    need = int(round(sum(raw.values()) + 1e-9)) - sum(units.values())
    for k in sorted(raw, key=lambda k: -(raw[k] - units[k]))[:max(need, 0)]:
        units[k] += 1
    for k, u in units.items():
        if 0 < u < min_units:
            units[k] = min_units if raw[k] >= min_units / 2 else 0
    return units


def _solve_person(units, months_of, rules: Rules, min_units: int):
    cells = [c for c, u in units.items() if u > 0]
    allm = sorted({t for c in cells for t in months_of[c]})
    cap_units = int(math.floor(rules.cap / rules.step + 1e-9))
    mdl = cp_model.CpModel(); x = {}
    for c in cells:
        for t in months_of[c]:
            x[c, t] = mdl.NewIntVar(0, units[c], "")
            if min_units > 1:                      # a month is either 0 or >= min_units
                y = mdl.NewBoolVar("")
                mdl.Add(x[c, t] >= min_units).OnlyEnforceIf(y); mdl.Add(x[c, t] == 0).OnlyEnforceIf(y.Not())
        mdl.Add(sum(x[c, t] for t in months_of[c]) == units[c])
    load = {t: sum(x[c, t] for c in cells if (c, t) in x) for t in allm}
    if rules.hard_cap:
        for t in allm: mdl.Add(load[t] <= cap_units)
    s = cp_model.CpSolver(); s.parameters.max_time_in_seconds = rules.seconds; s.parameters.num_workers = rules.workers

    if rules.objective == "level":
        peak = mdl.NewIntVar(0, max(sum(units.values()), 1), "")
        for t in allm: mdl.Add(load[t] <= peak)
        mdl.Minimize(peak)
        r = s.Solve(mdl)
        if r not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            return None, "infeasible" if r == cp_model.INFEASIBLE else "not solved"
        mdl.Add(peak <= int(s.ObjectiveValue()))
    mdl.ClearObjective()
    if rules.objective == "front":
        order = {t: i for i, t in enumerate(allm)}
        mdl.Minimize(sum(order[t] * x[c, t] for (c, t) in x))
    else:
        devs = []
        for c in cells:
            n = len(months_of[c])
            for t in months_of[c]:
                d = mdl.NewIntVar(0, units[c] * n, "")
                mdl.Add(d >= n * x[c, t] - units[c]); mdl.Add(d >= units[c] - n * x[c, t]); devs.append(d)
        mdl.Minimize(sum(devs))
    r2 = s.Solve(mdl)
    if r2 in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return {k: s.Value(v) for k, v in x.items()}, "ok" if r2 == cp_model.OPTIMAL else "ok (time limit)"
    return None, "infeasible" if r2 == cp_model.INFEASIBLE else "not solved"


def infeasibility_explanation(units: dict, months_of: dict, rules: Rules) -> str:
    """Describe a hard-cap capacity bottleneck that makes one person's schedule impossible."""
    if not rules.hard_cap:
        return "The solver found no feasible schedule, but no hard-cap capacity bottleneck applies."

    cells = [cell for cell, units_for_cell in units.items() if units_for_cell > 0]
    months = sorted({month for cell in cells for month in months_of[cell]})
    cap_units = int(math.floor(rules.cap / rules.step + 1e-9))
    bottlenecks = []
    for start in range(len(months)):
        for end in range(start, len(months)):
            window = set(months[start:end + 1])
            constrained = [cell for cell in cells if set(months_of[cell]).issubset(window)]
            required = sum(units[cell] for cell in constrained)
            capacity = len(window) * cap_units
            if required > capacity:
                bottlenecks.append((required - capacity, required, capacity, start, end, constrained))

    if not bottlenecks:
        return "The solver found no feasible schedule under the current rules; no single capacity-window bottleneck was identified."

    _, required, capacity, start, end, constrained = max(bottlenecks)
    first, last = months[start], months[end]
    period = f"{first:%b-%y}" if first == last else f"{first:%b-%y}–{last:%b-%y}"
    cells_text = ", ".join(f"{budget} / {milestone}" for budget, milestone in constrained[:3])
    if len(constrained) > 3:
        cells_text += f", and {len(constrained) - 3} more"
    return (
        f"Hard-cap bottleneck in {period}: {required * rules.step:g} d must be billed by this person "
        f"within those months, but the {rules.cap:g}-d/month cap allows only {capacity * rules.step:g} d. "
        f"Affected milestones: {cells_text}."
    )


@dataclass
class Result:
    schedule: pd.DataFrame
    check: pd.DataFrame
    flags: pd.DataFrame
    status: dict = field(default_factory=dict)
    rules: Rules | None = None


def run_optimization(loe: pd.DataFrame, ms: pd.DataFrame, rules: Rules, people: Iterable[str] | None = None,
                     progress: Callable[[int, int, str], None] | None = None) -> Result:
    d = clean_loe(loe)
    source_months = active_map(ms)
    act = active_map(ms, (9,) if rules.exclude_september else ())
    if people is not None: d = d[d.person.isin(list(people))]
    min_units = max(1, int(round(rules.min_bill / rules.step)))
    sched, check, flags, status = [], [], [], {}
    persons = sorted(d.person.unique())
    for i, p in enumerate(persons):
        if progress: progress(i, len(persons), p)
        sub = d[d.person == p]
        cells = {(r.budget, r.milestone): r.loe_days for r in sub.itertuples()}
        ok = {c: v for c, v in cells.items() if act.get(c)}
        for c, v in cells.items():
            if c not in ok:
                if c not in source_months:
                    why = "milestone not in table"
                elif not source_months[c]:
                    why = "milestone has no valid months (dates missing or end < start)"
                elif rules.exclude_september:
                    why = "milestone has no billable months after excluding September"
                else:
                    why = "milestone has no billable months"
                flags.append((p, c[0], c[1], "NOT SCHEDULED", f"{v:g} d allotted but {why}"))
        units = round_units(ok, rules.step, min_units)
        for c, v in ok.items():
            diff = units[c] * rules.step - v
            if abs(diff) > 1e-6:
                tag = "ROUNDED TO ZERO" if units[c] == 0 else "ROUNDED"
                flags.append((p, c[0], c[1], tag, f"allotted {v:.3f} d is not schedulable at a {rules.step:g}-day increment / {rules.min_bill:g}-day minimum; scheduled {units[c]*rules.step:g} d ({diff:+.3f})"))
        sol, st = _solve_person(units, act, rules, min_units) if any(units.values()) else ({}, "ok")
        status[p] = st
        if sol is None:
            if st == "infeasible":
                flags.append((p, "", "", "INFEASIBLE", infeasibility_explanation(units, act, rules)))
            else:
                flags.append((p, "", "", "NOT SOLVED", "The solver reached its time limit before finding a schedule. Increase solver time and try again."))
            for c, v in cells.items(): check.append((p, c[0], c[1], v, 0.0, -v, "Not scheduled"))
            continue
        got = {}
        for (c, t), u in sol.items():
            if u:
                sched.append((p, c[0], c[1], t, u * rules.step)); got[c] = got.get(c, 0) + u * rules.step
        for c, v in cells.items():
            g = got.get(c, 0.0)
            check.append((p, c[0], c[1], v, g, g - v, "OK" if abs(g - v) < 1e-6 else ("Not scheduled" if c not in ok else "Rounded")))
    if progress: progress(len(persons), len(persons), "")
    sch = pd.DataFrame(sched, columns=["person", "budget", "milestone", "month", "days"])
    ck = pd.DataFrame(check, columns=["person", "budget", "milestone", "allotted", "scheduled", "difference", "status"])
    if len(sch):
        load = sch.groupby(["person", "month"])["days"].sum()
        for (p, t), v in load.items():
            if v > rules.cap + 1e-9:
                flags.append((p, "", "", "OVER CAP", f"{t:%b-%y}: {v:.1f} d scheduled vs cap {rules.cap:g}"))
    fl = pd.DataFrame(flags, columns=["person", "budget", "milestone", "flag", "detail"])
    return Result(sch, ck, fl, status, rules)


def person_month(res: Result) -> pd.DataFrame:
    if res.schedule.empty: return pd.DataFrame()
    return res.schedule.pivot_table(index="person", columns="month", values="days", aggfunc="sum", fill_value=0.0)


def person_country_month(res: Result) -> pd.DataFrame:
    """Roll scheduled days up to each person within each country/budget and month."""
    if res.schedule.empty: return pd.DataFrame()
    return res.schedule.pivot_table(index=["budget", "person"], columns="month", values="days", aggfunc="sum", fill_value=0.0)


def to_workbook(res: Result, loe: pd.DataFrame, ms: pd.DataFrame) -> bytes:
    """Excel export: results plus the inputs that produced them."""
    pm = person_month(res).copy()
    if len(pm):
        pm.columns = [c.strftime("%b-%y") for c in pm.columns]
        pm["Total"] = pm.sum(axis=1)
    pcm = person_country_month(res).copy()
    if len(pcm):
        pcm.columns = [c.strftime("%b-%y") for c in pcm.columns]
        pcm["Total"] = pcm.sum(axis=1)
    summ = pd.DataFrame({"Metric": ["Billing increment (d)", "Minimum monthly entry (d)", "Monthly cap (d)", "Hard cap", "Exclude September", "Objective",
                                    "Cells", "Rounded", "Not scheduled", "Over-cap person-months", "Allotted days", "Scheduled days"],
                         "Value": [res.rules.step, res.rules.min_bill, res.rules.cap, res.rules.hard_cap, res.rules.exclude_september, OBJECTIVES[res.rules.objective],
                                   len(res.check), int(res.flags.flag.isin(["ROUNDED", "ROUNDED TO ZERO"]).sum()),
                                   int((res.flags.flag == "NOT SCHEDULED").sum()), int((res.flags.flag == "OVER CAP").sum()),
                                   round(res.check.allotted.sum(), 2), round(res.check.scheduled.sum(), 2)]})
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl", datetime_format="mmm-yy") as xw:
        summ.to_excel(xw, sheet_name="Summary", index=False)
        res.schedule.to_excel(xw, sheet_name="Schedule", index=False)
        pm.to_excel(xw, sheet_name="Person x Month")
        pcm.to_excel(xw, sheet_name="People by Country x Month")
        res.check.to_excel(xw, sheet_name="Check", index=False)
        res.flags.to_excel(xw, sheet_name="Flags", index=False)
        ms.to_excel(xw, sheet_name="Input - Milestones", index=False)
        loe.to_excel(xw, sheet_name="Input - LOE", index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(60, max(10, max(len(str(c.value or "")) for c in col[:200]) + 2))
    return buf.getvalue()


def inputs_workbook(loe: pd.DataFrame, ms: pd.DataFrame) -> bytes:
    """Template other people can fill in and send back (or upload into the app)."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl", datetime_format="mmm-yy", date_format="mmm-yy") as xw:
        ms.to_excel(xw, sheet_name="Milestones", index=False)
        loe.to_excel(xw, sheet_name="LOE", index=False)
        pd.DataFrame({"Notes": [
            "Milestones: start/end are calendar months (any day in the month works). Leave blank if unknown.",
            "LOE: one row per person x budget x milestone, loe_days = total days allotted over the milestone.",
            "To add a person or a milestone, add a row. Budget + milestone names must match between the two sheets."]}).to_excel(xw, sheet_name="README", index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(70, max(10, max(len(str(c.value or "")) for c in col[:200]) + 2))
    return buf.getvalue()


def read_inputs(file) -> tuple[pd.DataFrame, pd.DataFrame]:
    ms = pd.read_excel(file, "Milestones"); file.seek(0)
    loe = pd.read_excel(file, "LOE")
    for c in ("start", "end"): ms[c] = pd.to_datetime(ms[c], errors="coerce").dt.date
    return loe, ms
