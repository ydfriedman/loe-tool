# LOE tool

Plan how Palladium staff LOE (level of effort, in days) is billed to milestones month by month.

| Folder | What it is |
|---|---|
| `app/` | Streamlit app: milestone timeline, editable LOE / dates / business rules, optimizer, results. See `app/README.md`. |
| `scripts/` | Command-line versions of the pipeline (extract LOE from budgets, build milestone valid-months, schedule). |
| `out/` | Generated workbooks from the 13 Oct 2026 budgets and Oct 2026 workplans. |

```
pip install -r app/requirements.txt
streamlit run app/streamlit_app.py
```

Pipeline from raw files (scripts need `openpyxl`, `ortools`):

```
python3 scripts/extract_loe.py BUDGET_DIR out/LOE_by_person_milestone.xlsx
python3 scripts/build_milestone_months.py WORKPLAN_DIR out/Milestone_Valid_Months.xlsx
python3 scripts/schedule_loe.py out/LOE_by_person_milestone.xlsx out/Milestone_Valid_Months.xlsx out/LOE_monthly_schedule.xlsx
```

Notes: TBD placeholder staff are excluded; Word-workplan months in `build_milestone_months.py` are manually reviewed
judgements (with evidence recorded per row). Honduras and Sierra Leone have no workplan yet.
