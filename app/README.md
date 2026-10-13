# LOE scheduler (Streamlit)

Shows when milestones are active, who bills how many days to which milestone each month, lets you
change dates / days / business rules, and re-optimizes.

```
pip install -r app/requirements.txt
streamlit run app/streamlit_app.py
```

* **1 Milestone timeline** – Gantt per budget; edit start/end months (notes show where source workplans disagreed).
* **2 LOE per person** – edit days per person x budget x milestone; add rows for new people.
* **3 Check inputs** – missing dates, LOE on unknown milestones, milestones with no LOE, rounding preview.
* **4 Optimize & results** – heatmap vs cap, person / milestone views, flags, Excel export of results + inputs.
* **Sidebar** – billing increment, minimum monthly entry, monthly cap (flag or hard), objective, solver time;
  download the inputs workbook for others to fill in, upload it back.

Data lives in `app/data/` (seed CSVs from the 13 Oct budgets and the Oct workplans). Edits are held per browser
session; use "Download inputs" to keep them. The optimizer is in `app/scheduler.py` (no Streamlit dependency).
