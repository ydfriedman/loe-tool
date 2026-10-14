"""LOE scheduler: see when milestones are active, edit inputs and rules, and optimize who bills what, when."""
import datetime as dt
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

import scheduler as sc

DATA = Path(__file__).parent / "data"
st.set_page_config(page_title="LOE scheduler", page_icon="📅", layout="wide")


# ----------------------------------------------------------------------------- state
def seed():
    loe = pd.read_csv(DATA / "loe_seed.csv")
    ms = pd.read_csv(DATA / "milestones_seed.csv", parse_dates=["start", "end"])
    for c in ("start", "end"): ms[c] = ms[c].dt.date
    ms["notes"] = ms["notes"].fillna(""); ms["confidence"] = ms["confidence"].fillna("")
    return loe, ms


if "loe" not in st.session_state:
    st.session_state.loe, st.session_state.ms = seed()
    st.session_state.result = None
    st.session_state.rev = 0           # bumps when inputs are replaced, to reset the editors


def merge_back(full, edited, mask):
    """Write an edited subset back into the full table."""
    return pd.concat([full[~mask], edited], ignore_index=True)


# ----------------------------------------------------------------------------- sidebar: rules
sb = st.sidebar
sb.header("Business rules")
step = sb.selectbox("Billing increment (days)", [0.25, 0.5, 1.0], index=1,
                    help="Every monthly entry is a whole multiple of this.")
min_bill = sb.number_input("Smallest monthly entry when billing (days)", min_value=step, value=max(step, 0.5), step=step,
                           help="A person either bills 0 on a milestone in a month, or at least this much.")
cap = sb.number_input("Monthly cap per person (days)", min_value=1.0, max_value=31.0, value=21.7, step=0.1,
                      help="21.7 = 260 working days / 12.")
hard = sb.checkbox("Enforce cap as a hard limit", value=False,
                   help="Off: months over the cap are flagged. On: the solver must stay under it, and reports people for whom that is impossible.")
exclude_september = sb.checkbox("Do not bill in September", value=False,
                                help="Excludes every September from scheduling. The solver shifts that LOE into the remaining active months, such as October.")
objective_label, objective_info = sb.columns([4, 1])
objective_label.markdown("**Objective**")
with objective_info:
    with st.popover("Info", help="Explain the scheduling objectives"):
        st.markdown("""
        **Level** minimizes each person's busiest month, then spreads each milestone as evenly as possible.

        **Spread** spreads each milestone as evenly as possible without trying to reduce a person's overall monthly peaks.

        **Front** puts work in the earliest available months. If the hard cap is enabled, it fills earlier months up to that cap first.
        """)
objective = sb.radio("Objective", list(sc.OBJECTIVES), format_func=sc.OBJECTIVES.get, label_visibility="collapsed")
seconds = sb.slider("Solver time per person (seconds)", 1, 30, 5)
rules = sc.Rules(step=step, min_bill=min_bill, cap=cap, hard_cap=hard, exclude_september=exclude_september,
                 objective=objective, seconds=float(seconds))
for m in rules.validate(): sb.error(m)

sb.divider()
sb.header("Inputs")
sb.caption("Download the inputs, have others fill them in, then upload the file back.")
sb.download_button("Download inputs (.xlsx)", sc.inputs_workbook(st.session_state.loe, st.session_state.ms),
                   "loe_inputs.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
up = sb.file_uploader("Upload filled-in inputs", type="xlsx")
if up is not None and st.session_state.get("last_upload") != up.file_id:
    try:
        loe_u, ms_u = sc.read_inputs(up)
        st.session_state.loe, st.session_state.ms = loe_u, ms_u
        st.session_state.last_upload = up.file_id; st.session_state.result = None; st.session_state.rev += 1
        sb.success(f"Loaded {len(loe_u)} LOE rows and {len(ms_u)} milestones.")
        st.rerun()
    except Exception as e:  # noqa: BLE001
        sb.error(f"Could not read the file: {e}")
if sb.button("Reset to the original data"):
    st.session_state.loe, st.session_state.ms = seed()
    st.session_state.result = None; st.session_state.rev += 1; st.rerun()

st.title("LOE scheduler")
st.caption("Who bills how many days to which milestone, in which month. Edit the dates, days or rules, then optimize.")
tabs = st.tabs(["1 · Milestone timeline", "2 · LOE per person", "3 · Check inputs", "4 · Optimize & results"])
rev = st.session_state.rev

# ----------------------------------------------------------------------------- tab 1: timeline
with tabs[0]:
    ms = st.session_state.ms
    budgets = sorted(ms.budget.dropna().unique())
    c1, c2 = st.columns([3, 1])
    pick = c1.multiselect("Budgets", budgets, default=budgets[:3], placeholder="Choose one or more budgets")
    show_mob = c2.checkbox("Show mobilization rows", value=True)
    view = ms[ms.budget.isin(pick)].dropna(subset=["start", "end"]).copy()
    if not show_mob: view = view[~view.type.astype(str).str.startswith("Mobilization")]
    if view.empty:
        st.info("Pick at least one budget with dated milestones.")
    else:
        view["label"] = view.budget + " · " + view.milestone
        view["end_excl"] = pd.to_datetime(view.end) + pd.offsets.MonthBegin(1)
        view["start"] = pd.to_datetime(view.start)
        view["months"] = view.apply(lambda r: len(sc.months_between(r.start, r.end)), axis=1)
        chart = (alt.Chart(view).mark_bar(cornerRadius=3, height=16)
                 .encode(x=alt.X("start:T", title=None, axis=alt.Axis(format="%b-%y", tickCount="month")), x2="end_excl:T",
                         y=alt.Y("label:N", sort=None, title=None), color=alt.Color("budget:N", legend=alt.Legend(title="Budget")),
                         tooltip=["budget", "milestone", "title", alt.Tooltip("start:T", format="%b-%y"),
                                  alt.Tooltip("end:T", format="%b-%y"), "months", "confidence", "notes"])
                 .properties(height=max(120, 24 * len(view))))
        st.altair_chart(chart, width="stretch")
    st.subheader("Edit milestone dates")
    st.caption("Change start / end (any day in the month). Add rows for new milestones. Notes show where the source workplans disagreed.")
    sel = st.selectbox("Budget to edit", ["(all)"] + budgets, key=f"ms_sel{rev}")
    mask = (ms.budget == sel) if sel != "(all)" else pd.Series(True, index=ms.index)
    edited = st.data_editor(
        ms[mask], num_rows="dynamic", width="stretch", hide_index=True, key=f"ms_ed_{sel}_{rev}",
        column_config={"start": st.column_config.DateColumn("start", format="MMM YYYY"),
                       "end": st.column_config.DateColumn("end", format="MMM YYYY"),
                       "type": st.column_config.SelectboxColumn("type", options=["Milestone", "Mobilization (not in budget)"]),
                       "notes": st.column_config.TextColumn("notes", width="large"),
                       "title": st.column_config.TextColumn("title", width="large")})
    if sel != "(all)": edited["budget"] = edited["budget"].fillna(sel)
    st.session_state.ms = merge_back(ms, edited, mask) if sel != "(all)" else edited.reset_index(drop=True)

# ----------------------------------------------------------------------------- tab 2: LOE
with tabs[1]:
    loe = st.session_state.loe
    c1, c2 = st.columns(2)
    b_sel = c1.selectbox("Budget", ["(all)"] + sorted(loe.budget.dropna().unique()), key=f"loe_b{rev}")
    p_sel = c2.selectbox("Person", ["(all)"] + sorted(loe.person.dropna().unique()), key=f"loe_p{rev}")
    mask = pd.Series(True, index=loe.index)
    if b_sel != "(all)": mask &= loe.budget == b_sel
    if p_sel != "(all)": mask &= loe.person == p_sel
    st.caption(f"{int(mask.sum())} rows · {loe[mask].loe_days.sum():,.1f} days. Edit days, or add a row (person, budget, milestone, days).")
    known = sorted(st.session_state.ms.milestone.dropna().unique())
    edited = st.data_editor(
        loe[mask], num_rows="dynamic", width="stretch", hide_index=True, key=f"loe_ed_{b_sel}_{p_sel}_{rev}",
        column_config={"loe_days": st.column_config.NumberColumn("loe_days", min_value=0.0, step=0.5, format="%.2f"),
                       "milestone": st.column_config.SelectboxColumn("milestone", options=known),
                       "role": st.column_config.TextColumn("role", width="large")})
    if b_sel != "(all)": edited["budget"] = edited["budget"].fillna(b_sel)
    if p_sel != "(all)": edited["person"] = edited["person"].fillna(p_sel)
    st.session_state.loe = merge_back(loe, edited, mask).reset_index(drop=True)

    st.subheader("Totals by person")
    tot = sc.clean_loe(st.session_state.loe).groupby("person", as_index=False).agg(days=("loe_days", "sum"), budgets=("budget", "nunique"), milestones=("milestone", "count"))
    st.dataframe(tot.sort_values("days", ascending=False), width="stretch", hide_index=True,
                 column_config={"days": st.column_config.ProgressColumn("total days", min_value=0, max_value=260, format="%.1f")})

# ----------------------------------------------------------------------------- tab 3: checks
with tabs[2]:
    chk = sc.input_checks(st.session_state.loe, st.session_state.ms)
    if chk.empty:
        st.success("No input problems found.")
    else:
        st.warning(f"{len(chk)} things to fix or confirm before relying on the schedule.")
        kinds = chk.issue.value_counts()
        st.write(", ".join(f"**{k}**: {v}" for k, v in kinds.items()))
        st.dataframe(chk, width="stretch", hide_index=True)
    st.subheader("Rule sanity check")
    d = sc.clean_loe(st.session_state.loe)
    nonmult = d[(d.loe_days / step - (d.loe_days / step).round()).abs() > 1e-6]
    st.write(f"{len(nonmult)} of {len(d)} person-milestone allotments are not a multiple of {step:g} d and will be rounded "
             f"(the person's total moves by less than {step:g} d).")
    tiny = d[d.loe_days < min_bill / 2]
    st.write(f"{len(tiny)} allotments are under {min_bill/2:g} d and will round to zero.")

# ----------------------------------------------------------------------------- tab 4: optimize
with tabs[3]:
    d = sc.clean_loe(st.session_state.loe)
    c1, c2 = st.columns(2)
    scope_b = c1.multiselect("Only these budgets (blank = all)", sorted(d.budget.unique()))
    scope_p = c2.multiselect("Only these people (blank = all)", sorted(d.person.unique()))
    keep = pd.Series(True, index=d.index)
    if scope_b: keep &= d.budget.isin(scope_b)
    if scope_p: keep &= d.person.isin(scope_p)
    pool = d[keep]
    people = sorted(pool.person.unique())
    st.write(f"Will schedule **{len(people)}** people, **{len(pool)}** allotments, **{pool.loe_days.sum():,.1f}** days.")
    if st.button("Optimize", type="primary", disabled=bool(rules.validate()) or not people):
        bar = st.progress(0.0, text="Starting…")
        # scope by budget: restrict LOE rows, keep everyone's other rows out so loads reflect the scope
        res = sc.run_optimization(pool, st.session_state.ms, rules, people,
                                  progress=lambda i, n, p: bar.progress(i / max(n, 1), text=f"Solving {p}" if p else "Done"))
        bar.empty(); st.session_state.result = res; st.session_state.result_loe = pool.copy()
    res = st.session_state.result
    if res is None:
        st.info("Set the rules in the sidebar, then click Optimize.")
    else:
        r = res.rules
        st.caption(
            f"Results use: increment {r.step:g} d · min entry {r.min_bill:g} d · cap {r.cap:g} d "
            f"({'hard' if r.hard_cap else 'flag only'}) · "
            f"{'September excluded · ' if r.exclude_september else ''}{sc.OBJECTIVES[r.objective]}"
        )
        f = res.flags
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Allotted days", f"{res.check.allotted.sum():,.1f}")
        k2.metric("Scheduled days", f"{res.check.scheduled.sum():,.1f}")
        k3.metric("Rounded cells", int(f.flag.isin(["ROUNDED", "ROUNDED TO ZERO"]).sum()))
        k4.metric("Not scheduled", int((f.flag.isin(["NOT SCHEDULED", "INFEASIBLE", "NOT SOLVED"])).sum()))
        k5.metric("Over-cap person-months", int((f.flag == "OVER CAP").sum()))
        if (f.flag == "INFEASIBLE").any():
            st.error("Some people could not be scheduled under these rules. See the Flags tab for a person-specific capacity explanation.")
        if (f.flag == "NOT SOLVED").any():
            st.warning("The solver timed out before finding schedules for some people. See the Flags tab and increase solver time if needed.")
        sub = st.tabs(["Heatmap", "Person view", "Milestone view", "Full schedule", "Flags", "Check"])
        pm = sc.person_month(res)
        with sub[0]:
            if pm.empty: st.info("Nothing scheduled.")
            else:
                long = pm.reset_index().melt("person", var_name="month", value_name="days")
                long["month"] = pd.to_datetime(long["month"])
                order = pm.sum(axis=1).sort_values(ascending=False).index.tolist()
                base = alt.Chart(long).encode(x=alt.X("yearmonth(month):O", title=None, axis=alt.Axis(format="%b-%y", labelAngle=-45)),
                                              y=alt.Y("person:N", sort=order, title=None, axis=alt.Axis(labelOverlap=False, labelLimit=220)))
                heat = base.mark_rect().encode(color=alt.Color("days:Q", scale=alt.Scale(domain=[0, r.cap, r.cap * 2], range=["#f7fbff", "#6baed6", "#cb181d"], clamp=True),
                                                               legend=alt.Legend(title=f"days (cap {r.cap:g})")),
                                               tooltip=["person", alt.Tooltip("month:T", format="%b-%y"), alt.Tooltip("days:Q", format=".1f")])
                txt = base.mark_text(size=10).encode(text=alt.Text("days:Q", format=".0f"), color=alt.condition(alt.datum.days > r.cap, alt.value("white"), alt.value("#222")))
                st.altair_chart((heat + txt).properties(height=24 * len(order)), width="stretch")
                st.caption("Red = over the cap. Hover for exact days.")
        with sub[1]:
            who = st.selectbox("Person", sorted(res.schedule.person.unique()) if len(res.schedule) else [])
            if who:
                grouping = st.radio(
                    "Show monthly LOE by",
                    ["Budget × milestone", "Country / budget total"],
                    horizontal=True,
                )
                s = res.schedule[res.schedule.person == who].copy()
                if grouping == "Budget × milestone":
                    s["group"] = s.budget + " · " + s.milestone
                    group_title = "budget · milestone"
                else:
                    s["group"] = s.budget
                    group_title = "country / budget"
                chart_data = s.groupby(["month", "group"], as_index=False)["days"].sum()
                st.altair_chart(alt.Chart(chart_data).mark_bar().encode(
                    x=alt.X("yearmonth(month):O", title=None, axis=alt.Axis(format="%b-%y")), y=alt.Y("sum(days):Q", title="days"),
                    color=alt.Color("group:N", title=group_title),
                    tooltip=["group", alt.Tooltip("month:T", format="%b-%y"), "days"]).properties(height=320)
                    + alt.Chart(pd.DataFrame({"cap": [r.cap]})).mark_rule(color="red", strokeDash=[5, 3]).encode(y="cap:Q"), width="stretch")
                piv = s.pivot_table(index="group", columns=s.month.map(lambda x: x.strftime("%b-%y")), values="days", aggfunc="sum", fill_value=0)
                piv = piv[sorted(piv.columns, key=lambda c: dt.datetime.strptime(c, "%b-%y"))]
                piv["Total"] = piv.sum(axis=1)
                st.dataframe(piv, width="stretch")
        with sub[2]:
            if len(res.schedule):
                bm = res.schedule.assign(key=res.schedule.budget + " · " + res.schedule.milestone)
                opt = sorted(bm.key.unique())
                one = st.selectbox("Milestone", opt)
                s = bm[bm.key == one]
                st.altair_chart(alt.Chart(s).mark_bar().encode(
                    x=alt.X("yearmonth(month):O", title=None, axis=alt.Axis(format="%b-%y")), y=alt.Y("sum(days):Q", title="days"),
                    color=alt.Color("person:N"), tooltip=["person", alt.Tooltip("month:T", format="%b-%y"), "days"]).properties(height=300), width="stretch")
                piv = s.pivot_table(index="person", columns=s.month.map(lambda x: x.strftime("%b-%y")), values="days", aggfunc="sum", fill_value=0)
                piv = piv[sorted(piv.columns, key=lambda c: dt.datetime.strptime(c, "%b-%y"))]; piv["Total"] = piv.sum(axis=1)
                st.dataframe(piv, width="stretch")
        with sub[3]:
            st.dataframe(res.schedule.assign(month=res.schedule.month.map(lambda x: x.strftime("%b-%y"))), width="stretch", hide_index=True, height=480)
        with sub[4]:
            kinds = ["(all)"] + sorted(f.flag.unique())
            fk = st.selectbox("Flag type", kinds)
            st.dataframe(f if fk == "(all)" else f[f.flag == fk], width="stretch", hide_index=True, height=480)
        with sub[5]:
            st.dataframe(res.check, width="stretch", hide_index=True, height=480)
        st.download_button("Download results (.xlsx)", sc.to_workbook(res, st.session_state.result_loe, st.session_state.ms),
                           "loe_schedule.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
