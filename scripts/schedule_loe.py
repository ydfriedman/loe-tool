#!/usr/bin/env python3
"""Schedule each person's milestone LOE into months (0.5-day increments).

For every person x budget x milestone the allotted LOE is spread over the months
in which that milestone is active, in multiples of 0.5 day (a month is either 0 or
>= 0.5).  People are independent, so each is solved on its own with OR-tools CP-SAT:

  1. Rounding: allotments that are not multiples of 0.5 are rounded to 0.5 using
     largest-remainder rounding per person, so every cell moves by < 0.5 day and the
     person's total stays within 0.5 day of the true total.  Every change is flagged.
  2. Minimise the person's peak monthly load.
  3. With the peak held at that optimum, spread each milestone as evenly as possible
     across its active months (minimise absolute deviation from a flat profile).

Milestones with no active months (no workplan) cannot be scheduled and are flagged.
Months where a person exceeds --cap days are flagged (not forbidden).

Usage: python3 scripts/schedule_loe.py LOE_TABLE.xlsx MILESTONE_MONTHS.xlsx OUTPUT.xlsx [--cap 21.7]
"""
import argparse, datetime, math
from collections import defaultdict
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.table import Table, TableStyleInfo
from ortools.sat.python import cp_model

STEP = 0.5          # day increment; solver works in units of STEP
MILLI = 1000

def load(loe_path, mm_path):
    agg = defaultdict(float)
    for p, b, m, d, *_ in openpyxl.load_workbook(loe_path, data_only=True)["LOE Table"].iter_rows(min_row=2, values_only=True):
        agg[(p, b, str(m).strip())] += d
    active = defaultdict(list)
    for c, m, _, dt, *_ in openpyxl.load_workbook(mm_path, data_only=True)["Monthly (long)"].iter_rows(min_row=2, values_only=True):
        active[(c, m)].append(dt.date() if hasattr(dt, "date") else dt)
    return agg, {k: sorted(v) for k, v in active.items()}

def round_units(cells):
    """cells: {key: days}. Largest-remainder rounding to STEP, preserving the person total (nearest STEP)."""
    raw = {k: v / STEP for k, v in cells.items()}
    units = {k: math.floor(v + 1e-9) for k, v in raw.items()}
    need = int(round(sum(raw.values()) + 1e-9)) - sum(units.values())
    for k in sorted(raw, key=lambda k: -(raw[k] - units[k]))[:max(need, 0)]:
        units[k] += 1
    return units

def solve_person(units, months_of, tlimit):
    """units {cell: int}, months_of {cell: [months]} -> {(cell, month): units}, status"""
    cells = [c for c, u in units.items() if u > 0]
    allm = sorted({t for c in cells for t in months_of[c]})
    mdl = cp_model.CpModel()
    x = {}
    for c in cells:
        for t in months_of[c]:
            x[c, t] = mdl.NewIntVar(0, units[c], "")
        mdl.Add(sum(x[c, t] for t in months_of[c]) == units[c])
    load_t = {t: sum(x[c, t] for c in cells if (c, t) in x) for t in allm}
    peak = mdl.NewIntVar(0, sum(units.values()), "peak")
    for t in allm: mdl.Add(load_t[t] <= peak)
    mdl.Minimize(peak)
    s = cp_model.CpSolver(); s.parameters.max_time_in_seconds = tlimit; s.parameters.num_workers = 8
    r = s.Solve(mdl)
    if r not in (cp_model.OPTIMAL, cp_model.FEASIBLE): return None, "INFEASIBLE"
    best = int(s.ObjectiveValue()); st1 = "optimal" if r == cp_model.OPTIMAL else "feasible"
    # stage 2: hold peak, flatten each milestone
    mdl.Add(peak <= best); mdl.ClearObjective()
    devs = []
    for c in cells:
        n = len(months_of[c])
        for t in months_of[c]:
            d = mdl.NewIntVar(0, units[c] * n, "")
            mdl.Add(d >= n * x[c, t] - units[c]); mdl.Add(d >= units[c] - n * x[c, t]); devs.append(d)
    mdl.Minimize(sum(devs))
    r2 = s.Solve(mdl)
    if r2 in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return {k: s.Value(v) for k, v in x.items()}, f"peak {st1}; spread {'optimal' if r2 == cp_model.OPTIMAL else 'feasible'}"
    return None, "INFEASIBLE"

def main(a):
    agg, active = load(a.loe, a.months)
    by_person = defaultdict(dict)
    for (p, b, m), d in agg.items(): by_person[p][(b, m)] = d
    sched, flags, check, status = [], [], [], {}
    for p in sorted(by_person):
        cells = by_person[p]
        sched_cells = {c: d for c, d in cells.items() if active.get(c)}
        for c, d in cells.items():
            if not active.get(c):
                flags.append((p, c[0], c[1], "NOT SCHEDULED", f"{d:.2f} d allotted but milestone has no active months (no workplan)"))
        units = round_units(sched_cells)
        for c, d in sched_cells.items():
            diff = units[c] * STEP - d
            if abs(diff) > 1e-6:
                flags.append((p, c[0], c[1], "ROUNDED TO ZERO" if units[c] == 0 else "ROUNDED", f"allotted {d:.3f} d is not a multiple of {STEP}; scheduled {units[c]*STEP:.1f} d ({diff:+.3f})"))
        sol, st = solve_person(units, active, a.time)
        status[p] = st
        if sol is None:
            flags.append((p, "", "", "SOLVER", st)); continue
        for (c, t), u in sol.items():
            if u: sched.append((p, c[0], c[1], t, u * STEP))
        got = defaultdict(float)
        for (c, t), u in sol.items(): got[c] += u * STEP
        for c, d in cells.items():
            check.append((p, c[0], c[1], d, got.get(c, 0.0), got.get(c, 0.0) - d))
    # overload flags
    tot = defaultdict(float)
    for p, b, m, t, d in sched: tot[p, t] += d
    for (p, t), d in sorted(tot.items()):
        if d > a.cap + 1e-9:
            flags.append((p, "", "", "OVER CAP", f"{t:%b-%y}: {d:.1f} d scheduled vs cap {a.cap}"))

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Schedule"
    ws.append(["Person", "Budget", "Milestone", "Month", "Days"])
    for p, b, m, t, d in sorted(sched, key=lambda r: (r[0], r[1], r[2], r[3])): ws.append([p, b, m, t, d])
    n = ws.max_row
    for r in ws.iter_rows(min_row=2, min_col=4, max_col=4): r[0].number_format = "mmm-yy"
    tb = Table(displayName="Schedule", ref=f"A1:E{n}"); tb.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True); ws.add_table(tb)
    for col, w in zip("ABCDE", [28, 20, 12, 10, 8]): ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"

    months = sorted({t for *_, t, _ in sched}); people = sorted({r[0] for r in sched})
    pm = wb.create_sheet("Person x Month")
    pm.append(["Person"] + months + ["Total", "Peak month"])
    for c in range(2, len(months) + 2): pm.cell(1, c).number_format = "mmm-yy"
    for p in people:  # static values (no formulas) so previews and viewers show numbers
        row = [round(tot.get((p, t), 0), 2) for t in months]
        pm.append([p] + row + [round(sum(row), 2), max(row)])
    last = L(len(months) + 1)
    from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
    rng = f"B2:{last}{pm.max_row}"
    pm.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=[str(a.cap)], fill=PatternFill("solid", start_color="F4B6B6", end_color="F4B6B6"), font=Font(bold=True, color="9C0006")))
    pm.conditional_formatting.add(rng, ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF", end_type="num", end_value=a.cap, end_color="9DC3E6"))
    pm.cell(pm.max_row + 2, 1, f"Red = over the {a.cap}-day monthly cap. Blue shading scales from 0 to the cap.")
    pm.column_dimensions["A"].width = 28; pm.freeze_panes = "B2"

    ck = wb.create_sheet("Check")
    ck.append(["Person", "Budget", "Milestone", "Allotted (days)", "Scheduled (days)", "Difference", "Status"])
    for p, b, m, d, g, df in check:
        ck.append([p, b, m, round(d, 3), g, round(df, 3), "OK" if abs(df) < 1e-6 else ("Not scheduled" if g == 0 and not active.get((b, m)) else "Rounded to 0.5")])
    t2 = Table(displayName="Check", ref=f"A1:G{ck.max_row}"); t2.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True); ck.add_table(t2)
    for col, w in zip("ABCDEFG", [28, 20, 12, 15, 16, 11, 16]): ck.column_dimensions[col].width = w
    ck.freeze_panes = "A2"

    fl = wb.create_sheet("Flags"); fl.append(["Person", "Budget", "Milestone", "Flag", "Detail"])
    for f in sorted(flags, key=lambda r: (r[3], r[0], r[1], r[2])): fl.append(list(f))
    for col, w in zip("ABCDE", [28, 20, 12, 16, 110]): fl.column_dimensions[col].width = w
    fl.freeze_panes = "A2"

    sm = wb.create_sheet("Summary", 0)
    kinds = defaultdict(int)
    for f in flags: kinds[f[3]] += 1
    sm.append(["Metric", "Value"])
    sm.append(["People scheduled", len(people)])
    sm.append(["Person-milestone cells in LOE table", len(check)])
    sm.append(["Cells scheduled exactly as allotted", sum(1 for r in check if abs(r[5]) < 1e-6)])
    sm.append(["Cells rounded to a 0.5 multiple", kinds["ROUNDED"] + kinds["ROUNDED TO ZERO"]])
    sm.append(["  of which rounded down to 0 (allotment < 0.25 d)", kinds["ROUNDED TO ZERO"]])
    sm.append(["Cells not schedulable (no active months)", kinds["NOT SCHEDULED"]])
    sm.append(["Person-months over cap", kinds["OVER CAP"]])
    sm.append(["Cap used (days / month)", a.cap])
    sm.append(["Allotted days (all cells)", round(sum(r[3] for r in check), 2)])
    sm.append(["Scheduled days", round(sum(r[4] for r in check), 2)])
    sm.append(["Solver status", "; ".join(sorted(set(status.values())))])
    sm.column_dimensions["A"].width = 44; sm.column_dimensions["B"].width = 40
    wb.save(a.out)
    print(f"{len(sched)} schedule rows, {len(people)} people -> {a.out}")
    for r in sm.iter_rows(min_row=2, values_only=True): print(f"  {r[0]}: {r[1]}")

if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("loe"); ap.add_argument("months"); ap.add_argument("out")
    ap.add_argument("--cap", type=float, default=21.7, help="flag person-months above this many days (default 21.7 = 260/12)")
    ap.add_argument("--time", type=float, default=20, help="solver seconds per stage per person")
    main(ap.parse_args())
