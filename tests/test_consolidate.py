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
        wide, coa, entities, fx, ic, assoc = c.load_inputs(data_dir)
        usd, journal, elim, consolidated, em = c.consolidate(wide, coa, entities, fx, ic, assoc)
        self.em = em
        return c.summarise(usd, entities, consolidated), usd

    def copy_data(self):
        d = Path(tempfile.mkdtemp()); shutil.copytree(DATA, d, dirs_exist_ok=True)
        self.addCleanup(shutil.rmtree, d)
        return d

    def test_sample_group_consolidates_and_balances(self):
        s, usd = self.run_on(DATA)
        self.assertEqual(s["Revenue"], 25480)
        self.assertAlmostEqual(s["Goodwill"], 641.0, places=2)
        self.assertAlmostEqual(s["Profit for the year"], 3608.7, places=1)
        self.assertAlmostEqual(s["Attributable to NCI"], 239.78, places=2)
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

    def test_sub_subsidiary_uses_effective_interest(self):
        wide, coa, entities, fx, ic, assoc = c.load_inputs(DATA)
        self.assertAlmostEqual(entities.set_index("entity").effective["SUB3"], 0.7)
        _, journal, *_ = c.consolidate(wide, coa, entities, fx, ic, assoc)
        gw = sum(a for n, acct, a in journal if acct == c.GW and "Three" in n)
        # GBP 60 of goodwill (600 cost - 70% x GBP 600 net assets x 1.25) retranslated at the 1.27 closing rate
        self.assertAlmostEqual(gw, 60 * 1.27, places=6)

    def test_sub_subsidiary_held_by_foreign_partly_owned_subsidiary(self):
        wide, coa, entities, fx, ic, assoc = c.load_inputs(DATA)
        self.assertAlmostEqual(entities.set_index("entity").effective["SUB4"], 0.6)
        _, journal, *_ = c.consolidate(wide, coa, entities, fx, ic, assoc)
        sub4 = [(acct, a) for n, acct, a in journal if n == "Investment in Subsidiary Four (France)"]
        # EUR 450 investment held by Sub Two, removed at the 1.08 closing rate
        self.assertAlmostEqual(sum(a for acct, a in sub4 if acct == c.INV), -450 * 1.08, places=6)
        # Group share of goodwill: 80% x (EUR 450 - 75% x EUR 500) = EUR 60, at the 1.08 closing rate
        self.assertAlmostEqual(sum(a for acct, a in sub4 if acct == c.GW), 60 * 1.08, places=6)

    def test_indirect_holding_through_partly_owned_subsidiary(self):
        d = self.copy_data()
        e = pd.read_csv(d / "entities.csv")
        e.loc[e.entity == "SUB1", "ownership"] = 0.9
        e.to_csv(d / "entities.csv", index=False)
        wide, coa, entities, fx, ic, assoc = c.load_inputs(d)
        self.assertAlmostEqual(entities.set_index("entity").effective["SUB3"], 0.63)
        s, _ = self.run_on(d)
        self.assertAlmostEqual(s["Total assets"], s["Total equity and liabilities"], places=2)

    def test_unknown_holding_company_is_rejected(self):
        d = self.copy_data()
        e = pd.read_csv(d / "entities.csv")
        e.loc[e.entity == "SUB3", "held_by"] = "SUB9"
        e.to_csv(d / "entities.csv", index=False)
        with self.assertRaisesRegex(c.ValidationError, "held by SUB9"):
            self.run_on(d)

    def test_associates_use_equity_method(self):
        s, _ = self.run_on(DATA)
        em = self.em.set_index("entity")
        # Associate One: 900 cost + 122.53 opening share + 179.72 profit - 81.69 dividends - 12 unrealised profit
        self.assertAlmostEqual(em.at["ASSOC1", "carrying_amount"], 1108.56, places=1)
        # Associate Two: 600 + 100 + 120 - 50 dividends - 40 impairment
        self.assertAlmostEqual(em.at["ASSOC2", "carrying_amount"], 730, places=2)
        self.assertAlmostEqual(s["Investment in associates"], em.carrying_amount.sum(), places=2)
        self.assertAlmostEqual(s["Share of profit of associates"], 179.72 + 120 - 40, places=1)

    def test_dividend_mismatch_with_associates_is_reported(self):
        d = self.copy_data()
        a = pd.read_csv(d / "associates.csv")
        a.loc[1, "dividends_paid"] = 300
        a.to_csv(d / "associates.csv", index=False)
        with self.assertRaisesRegex(c.ValidationError, "Dividend income from associates"):
            self.run_on(d)

    def test_associate_over_50_percent_is_rejected(self):
        d = self.copy_data()
        a = pd.read_csv(d / "associates.csv")
        a.loc[0, "ownership"] = 0.6
        a.to_csv(d / "associates.csv", index=False)
        with self.assertRaisesRegex(c.ValidationError, "normally controlled"):
            self.run_on(d)

    def test_workbook_is_written(self):
        out = Path(tempfile.mkdtemp()) / "group.xlsx"
        self.addCleanup(shutil.rmtree, out.parent)
        self.assertEqual(c.main(["--data", str(DATA), "--out", str(out)]), 0)
        self.assertTrue(out.exists())


if __name__ == "__main__":
    unittest.main()
