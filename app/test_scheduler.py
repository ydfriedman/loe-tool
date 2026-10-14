import datetime as dt
import unittest

import pandas as pd

import scheduler as sc


class SchedulerTest(unittest.TestCase):
    def test_hard_cap_diagnosis_identifies_capacity_window(self):
        rules = sc.Rules(step=1.0, cap=1.0, hard_cap=True)
        units = {("Botswana", "BOT01"): 4}
        months = {("Botswana", "BOT01"): [dt.date(2026, 10, 1), dt.date(2026, 11, 1)]}

        explanation = sc.infeasibility_explanation(units, months, rules)

        self.assertIn("Oct-26–Nov-26", explanation)
        self.assertIn("4 d", explanation)
        self.assertIn("2 d", explanation)
        self.assertIn("Botswana / BOT01", explanation)

    def test_hard_cap_infeasibility_is_reported_for_person(self):
        loe = pd.DataFrame([{"person": "Ada", "budget": "Botswana", "milestone": "BOT01", "loe_days": 4.0}])
        milestones = pd.DataFrame([{
            "budget": "Botswana", "milestone": "BOT01", "start": dt.date(2026, 10, 1), "end": dt.date(2026, 11, 1),
        }])

        result = sc.run_optimization(loe, milestones, sc.Rules(step=1.0, cap=1.0, hard_cap=True))

        flag = result.flags.iloc[0]
        self.assertEqual("INFEASIBLE", flag.flag)
        self.assertIn("Hard-cap bottleneck", flag.detail)

    def test_person_country_month_rollup(self):
        result = sc.Result(
            schedule=pd.DataFrame([
                ("Ada", "Botswana", "BOT01", dt.date(2026, 10, 1), 2.0),
                ("Ada", "Botswana", "BOT02", dt.date(2026, 10, 1), 1.0),
                ("Ada", "Kenya", "KEN01", dt.date(2026, 10, 1), 4.0),
            ], columns=["person", "budget", "milestone", "month", "days"]),
            check=pd.DataFrame(),
            flags=pd.DataFrame(),
        )

        rollup = sc.person_country_month(result)

        self.assertEqual(3.0, rollup.loc[("Botswana", "Ada"), dt.date(2026, 10, 1)])
        self.assertEqual(4.0, rollup.loc[("Kenya", "Ada"), dt.date(2026, 10, 1)])

    def test_front_objective_bills_in_earliest_months(self):
        loe = pd.DataFrame([{"person": "Ada", "budget": "Botswana", "milestone": "BOT01", "loe_days": 4.0}])
        milestones = pd.DataFrame([{
            "budget": "Botswana", "milestone": "BOT01", "start": dt.date(2026, 10, 1), "end": dt.date(2026, 12, 1),
        }])

        result = sc.run_optimization(loe, milestones, sc.Rules(step=1.0, objective="front"))

        scheduled = result.schedule.set_index("month").days
        self.assertEqual(4.0, scheduled[dt.date(2026, 10, 1)])
        self.assertNotIn(dt.date(2026, 11, 1), scheduled.index)

    def test_excluding_september_shifts_loe_to_october(self):
        loe = pd.DataFrame([{"person": "Ada", "budget": "Botswana", "milestone": "BOT01", "loe_days": 4.0}])
        milestones = pd.DataFrame([{
            "budget": "Botswana", "milestone": "BOT01", "start": dt.date(2026, 9, 1), "end": dt.date(2026, 10, 1),
        }])

        result = sc.run_optimization(
            loe, milestones, sc.Rules(step=1.0, objective="front", exclude_september=True),
        )

        scheduled = result.schedule.set_index("month").days
        self.assertEqual(4.0, scheduled[dt.date(2026, 10, 1)])
        self.assertNotIn(dt.date(2026, 9, 1), scheduled.index)


if __name__ == "__main__":
    unittest.main()
