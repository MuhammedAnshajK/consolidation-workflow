"""Run: python -m unittest discover tests"""
import shutil
import tempfile
import unittest
from pathlib import Path

import pandas as pd

import consolidate as c

DATA = Path(__file__).resolve().parent.parent / "data"


class ConsolidateTest(unittest.TestCase):
    def run_on(self, data_dir):
        wide, coa, entities, fx, ic = c.load_inputs(data_dir)
        usd, journal, elim, consolidated = c.consolidate(wide, coa, entities, fx, ic)
        return c.summarise(usd, entities, consolidated), usd

    def copy_data(self):
        d = Path(tempfile.mkdtemp()); shutil.copytree(DATA, d, dirs_exist_ok=True)
        self.addCleanup(shutil.rmtree, d)
        return d

    def test_sample_group_consolidates_and_balances(self):
        s, usd = self.run_on(DATA)
        self.assertEqual(s["Revenue"], 22124)
        self.assertEqual(s["Goodwill"], 500)
        self.assertAlmostEqual(s["Profit for the year"], 2928.7, places=1)
        self.assertAlmostEqual(s["Attributable to NCI"], 93.74, places=2)
        self.assertAlmostEqual(s["Total assets"], s["Total equity and liabilities"], places=2)

    def test_translated_trial_balance_balances(self):
        _, usd = self.run_on(DATA)
        for entity in usd.columns:
            self.assertAlmostEqual(usd[entity].sum(), 0, places=6)

    def test_intercompany_mismatch_is_reported(self):
        d = self.copy_data()
        ic = pd.read_csv(d / "intercompany.csv")
        ic.loc[ic.type == "balance", "amount_group_currency"] += 25
        ic.to_csv(d / "intercompany.csv", index=False)
        with self.assertRaisesRegex(c.ValidationError, "Intercompany receivables do not eliminate"):
            self.run_on(d)

    def test_unbalanced_trial_balance_is_rejected(self):
        d = self.copy_data()
        tb = pd.read_csv(d / "trial_balances.csv")
        tb.loc[0, "debit"] += 10
        tb.to_csv(d / "trial_balances.csv", index=False)
        with self.assertRaisesRegex(c.ValidationError, "PARENT does not balance"):
            self.run_on(d)

    def test_workbook_is_written(self):
        out = Path(tempfile.mkdtemp()) / "group.xlsx"
        self.addCleanup(shutil.rmtree, out.parent)
        self.assertEqual(c.main(["--data", str(DATA), "--out", str(out)]), 0)
        self.assertTrue(out.exists())


if __name__ == "__main__":
    unittest.main()
