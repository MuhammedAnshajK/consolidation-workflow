"""Consolidates a parent and its subsidiaries into group financial statements.

Steps: load each entity's trial balance, translate foreign subsidiaries into the
group currency, eliminate intercompany balances, sales, unrealised profit and the
investment in subsidiaries, recognise goodwill and non-controlling interests, and
write an Excel workbook in which every step is a formula.

Run:  python consolidate.py
"""
import argparse
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

GROUP = "Illustrative Group"
GROUP_CCY = "USD"
UNIT = f"{GROUP_CCY} thousands"
TOLERANCE = 0.5  # largest intercompany mismatch accepted, in group currency

# Accounts the consolidation adds on top of the entity chart of accounts.
GROUP_ACCOUNTS = [(1700, "Goodwill", "A"), (3200, "Translation reserve", "E"), (3300, "Non-controlling interests", "E")]
SC, RE, INV, GW, CTA, NCI = 3000, 3100, 1600, 1700, 3200, 3300
IC_REC, IC_PAY, IC_REV, IC_PUR, COS, INVENTORY = 1150, 2150, 4050, 5050, 5000, 1200


class ValidationError(Exception):
    pass


# ---------------------------------------------------------------- load and validate
def load_inputs(data_dir):
    d = Path(data_dir)
    tb = pd.read_csv(d / "trial_balances.csv")
    coa = pd.read_csv(d / "chart_of_accounts.csv")
    entities = pd.read_csv(d / "entities.csv")
    fx = pd.read_csv(d / "fx_rates.csv").set_index("currency")
    ic = pd.read_csv(d / "intercompany.csv")

    tb["balance"] = tb.debit.fillna(0) - tb.credit.fillna(0)
    for ent, rows in tb.groupby("entity"):
        if round(rows.balance.sum(), 2) != 0:
            raise ValidationError(f"Trial balance for {ent} does not balance (difference {rows.balance.sum():,.2f})")
    unknown = sorted(set(tb.account_code) - set(coa.account_code))
    if unknown:
        raise ValidationError(f"Accounts not in the chart of accounts: {unknown}")
    missing = sorted(set(tb.entity) - set(entities.entity))
    if missing:
        raise ValidationError(f"Entities missing from entities.csv: {missing}")
    no_rate = sorted(set(entities.currency) - set(fx.index))
    if no_rate:
        raise ValidationError(f"No exchange rates for: {no_rate}")
    if entities.iloc[0].ownership != 1:
        raise ValidationError("The first entity in entities.csv must be the parent (ownership 1)")

    wide = tb.pivot_table(index="account_code", columns="entity", values="balance", aggfunc="sum").fillna(0)
    wide = wide.reindex(coa.account_code).fillna(0)[list(entities.entity)]
    return wide, coa, entities, fx, ic


# ---------------------------------------------------------------- consolidation (pandas; mirrors the Excel formulas)
def translate(wide, coa, entities, fx):
    """Translates each entity into the group currency: assets and liabilities at the closing rate,
    income and expenses at the average rate, share capital at the historical rate, opening retained
    earnings at the group-currency amount carried forward. The difference goes to the translation reserve."""
    types = coa.set_index("account_code").type
    out = {}
    for _, e in entities.iterrows():
        local = wide[e.entity]
        if e.currency == GROUP_CCY:
            usd = local.copy()
        else:
            r = fx.loc[e.currency]
            rate = types.map({"A": r.closing, "L": r.closing, "I": r.average, "X": r.average, "E": r.historical})
            usd = local * rate
            usd[RE] = -float(e.opening_re_group_currency)
        usd[CTA] = -usd.sum()
        out[e.entity] = usd
    return pd.DataFrame(out).fillna(0)


def eliminations(usd, entities, ic):
    """Returns a list of (entry, description, account, amount) with debits positive."""
    j = []
    for _, e in entities.iloc[1:].iterrows():
        own, inv, pre = e.ownership, e.investment_cost, e.pre_acquisition_re
        sc = -usd.at[SC, e.entity]
        goodwill = inv - own * (sc + pre)
        n = f"Investment in {e['name']}"
        j += [(n, SC, sc), (n, RE, pre), (n, GW, goodwill), (n, INV, -inv), (n, NCI, -(1 - own) * (sc + pre))]
        if own < 1:
            n = f"NCI share of {e['name']} post-acquisition reserves"
            post_re = -usd.at[RE, e.entity] - pre
            cta = usd.at[CTA, e.entity]
            j += [(n, RE, (1 - own) * post_re), (n, NCI, -(1 - own) * post_re),
                  (n, CTA, -(1 - own) * cta), (n, NCI, (1 - own) * cta)]
    for _, r in ic.iterrows():
        if r.type == "balance":
            n = f"Intercompany balance: {r.description}"
            j += [(n, IC_PAY, r.amount_group_currency), (n, IC_REC, -r.amount_group_currency)]
        elif r.type == "sale":
            n = f"Intercompany sale: {r.description}"
            j += [(n, IC_REV, r.amount_group_currency), (n, IC_PUR, -r.amount_group_currency)]
        elif r.type == "unrealised_profit":
            n = f"Unrealised profit in inventory: {r.description}"
            j += [(n, COS, r.amount_group_currency), (n, INVENTORY, -r.amount_group_currency)]
    return j


def consolidate(wide, coa, entities, fx, ic):
    usd = translate(wide, coa, entities, fx)
    usd = usd.reindex(list(usd.index) + [a[0] for a in GROUP_ACCOUNTS if a[0] not in usd.index]).fillna(0)
    journal = eliminations(usd, entities, ic)
    elim = pd.Series(0.0, index=usd.index)
    for _, acct, amt in journal:
        elim[acct] += amt
    for entry in {j[0] for j in journal}:
        if abs(sum(a for n, _, a in journal if n == entry)) > 1e-6:
            raise ValidationError(f"Elimination entry does not balance: {entry}")
    consolidated = usd.sum(axis=1) + elim
    for acct, label in ((IC_REC, "receivables"), (IC_PAY, "payables"), (IC_REV, "revenue"), (IC_PUR, "purchases")):
        if abs(consolidated[acct]) > TOLERANCE:
            raise ValidationError(f"Intercompany {label} do not eliminate: {consolidated[acct]:,.2f} left over. "
                                  f"Check intercompany.csv against the entity trial balances.")
    return usd, journal, elim, consolidated


def summarise(usd, entities, consolidated):
    c = -consolidated
    profit = c[[a for a in c.index if 4000 <= a < 9000]].sum()
    nci_profit = sum((1 - e.ownership) * -usd[e.entity][[a for a in usd.index if 4000 <= a < 9000]].sum()
                     for _, e in entities.iloc[1:].iterrows())
    s = {"Revenue": c[4000] + c[4050], "Profit for the year": profit, "Attributable to owners": profit - nci_profit,
         "Attributable to NCI": nci_profit, "Goodwill": -c[GW],
         "Total assets": -c[[a for a in c.index if a < 2000]].sum(),
         "Equity attributable to owners": c[SC] + c[RE] + c[CTA] + profit - nci_profit,
         "Non-controlling interests": c[NCI] + nci_profit,
         "Total liabilities": c[[a for a in c.index if 2000 <= a < 3000]].sum()}
    s["Total equity and liabilities"] = s["Equity attributable to owners"] + s["Non-controlling interests"] + s["Total liabilities"]
    return {k: round(float(v), 2) for k, v in s.items()}


# ---------------------------------------------------------------- Excel output
FONT = "Arial"
HFILL = PatternFill("solid", fgColor="1F4E78")
TOP = Border(top=Side(style="thin"))
DBL = Border(top=Side(style="thin"), bottom=Side(style="double"))
NUM = '#,##0;(#,##0);"-"'
RATE = "0.0000"


def _f(bold=False, color="000000", italic=False, size=10):
    return Font(name=FONT, bold=bold, color=color, italic=italic, size=size)


def _sheet(wb, title, heading, sub=UNIT):
    ws = wb.create_sheet(title)
    ws["B1"] = GROUP; ws["B1"].font = _f(True, size=14)
    ws["B2"] = heading; ws["B2"].font = _f(True, size=12)
    ws["B3"] = sub; ws["B3"].font = _f(italic=True, color="595959")
    ws.column_dimensions["A"].width = 2
    ws.sheet_view.showGridLines = False
    return ws


def _header(ws, row, labels, first_col=2, widths=None):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=first_col + i, value=lab)
        c.font = _f(True, "FFFFFF"); c.fill = HFILL
        c.alignment = Alignment(horizontal="left" if i < 2 else "right", wrap_text=True)
        if widths: ws.column_dimensions[c.column_letter].width = widths[i]


def write_workbook(wide, coa, entities, fx, ic, journal, out_path):
    wb = Workbook(); wb.remove(wb.active)
    ents = list(entities.entity)
    accounts = list(coa.itertuples(index=False)) + [type(coa.iloc[0:1].itertuples(index=False).__next__())(*a)
                                                     for a in GROUP_ACCOUNTS]

    # Inputs: entities, FX, intercompany
    inp = _sheet(wb, "Inputs", "Group structure, exchange rates and intercompany data")
    _header(inp, 5, ["Entity", "Name", "Currency", "Ownership", "Investment cost", "Pre-acq. retained earnings",
                     "Opening RE (USD)"], widths=[10, 28, 10, 11, 12, 14, 12])
    ent_row = {}
    for i, e in enumerate(entities.itertuples(index=False)):
        r = 6 + i; ent_row[e.entity] = r
        for j, v in enumerate([e.entity, e.name, e.currency, e.ownership, e.investment_cost, e.pre_acquisition_re,
                               None if pd.isna(e.opening_re_group_currency) else e.opening_re_group_currency]):
            c = inp.cell(row=r, column=2 + j, value=v); c.font = _f(color="0000FF" if j >= 3 else "000000")
            c.number_format = "0%" if j == 3 else NUM
    fx_top = 6 + len(ents) + 2
    _header(inp, fx_top, ["Currency", "", "Closing", "Average", "Historical"])
    fx_row = {}
    for i, (ccy, r_) in enumerate(fx.iterrows()):
        r = fx_top + 1 + i; fx_row[ccy] = r
        inp.cell(row=r, column=2, value=ccy).font = _f()
        for j, k in enumerate(("closing", "average", "historical")):
            c = inp.cell(row=r, column=4 + j, value=float(r_[k])); c.font = _f(color="0000FF"); c.number_format = RATE
    ic_top = fx_top + len(fx) + 3
    _header(inp, ic_top, ["Type", "Description", "From", "To", "Amount (USD)"])
    inp.column_dimensions["C"].width = 46
    ic_row = []
    for i, x in enumerate(ic.itertuples(index=False)):
        r = ic_top + 1 + i; ic_row.append(r)
        for j, v in enumerate([x.type, x.description, x.from_entity, x.to_entity, x.amount_group_currency]):
            c = inp.cell(row=r, column=2 + j, value=v); c.font = _f(color="0000FF" if j == 4 else "000000")
            if j == 4: c.number_format = NUM

    # Trial balances in local currency
    tbs = _sheet(wb, "Trial Balances", "Entity trial balances in local currency (debit +, credit -)", "thousands of local currency")
    _header(tbs, 5, ["Code", "Account", "Type"] + [f"{e} ({entities.set_index('entity').currency[e]})" for e in ents],
            widths=[8, 30, 6] + [13] * len(ents))
    tb_row = {}
    for i, a in enumerate(coa.itertuples(index=False)):
        r = 6 + i; tb_row[a.account_code] = r
        for j, v in enumerate([a.account_code, a.account_name, a.type]):
            tbs.cell(row=r, column=2 + j, value=v).font = _f()
        for k, e in enumerate(ents):
            c = tbs.cell(row=r, column=5 + k, value=float(wide.at[a.account_code, e]))
            c.font = _f(color="0000FF"); c.number_format = NUM
    last_tb = 6 + len(coa) - 1
    tbs.cell(row=last_tb + 1, column=3, value="Check (should be zero)").font = _f(True)
    for k in range(len(ents)):
        col = tbs.cell(row=6, column=5 + k).column_letter
        c = tbs[f"{col}{last_tb + 1}"]; c.value = f"=SUM({col}6:{col}{last_tb})"; c.number_format = NUM
        c.font = _f(True); c.border = TOP

    # Translation into USD
    tr = _sheet(wb, "Translation", "Translation into USD")
    tr_cols = {}
    widths = [8, 30]
    labels = ["Code", "Account"]
    for e in ents:
        labels += [f"{e} local", f"{e} rate", f"{e} USD"]; widths += [11, 9, 11]
    _header(tr, 5, labels, widths=widths)
    tr_row = {}
    rows = [(a.account_code, a.account_name, a.type) for a in coa.itertuples(index=False)] + [(CTA, "Translation reserve", "E")]
    for i, (code, name, typ) in enumerate(rows):
        r = 6 + i; tr_row[code] = r
        tr.cell(row=r, column=2, value=code).font = _f(); tr.cell(row=r, column=3, value=name).font = _f()
    for k, e in enumerate(ents):
        lc, rc, uc = (tr.cell(row=5, column=4 + 3 * k + m).column_letter for m in range(3))
        tr_cols[e] = uc
        tbcol = tbs.cell(row=5, column=5 + k).column_letter
        fr = fx_row[entities.set_index("entity").currency[e]]
        for code, name, typ in rows[:-1]:
            r = tr_row[code]
            tr[f"{lc}{r}"] = f"='Trial Balances'!{tbcol}{tb_row[code]}"
            basis = {"A": "D", "L": "D", "I": "E", "X": "E", "E": "F"}[typ]
            tr[f"{rc}{r}"] = f"=Inputs!${basis}${fr}"
            if code == RE and entities.set_index("entity").currency[e] != GROUP_CCY:
                tr[f"{uc}{r}"] = f"=-Inputs!$H${ent_row[e]}"
                tr[f"{rc}{r}"] = f"=IF({lc}{r}=0,0,{uc}{r}/{lc}{r})"
            else:
                tr[f"{uc}{r}"] = f"={lc}{r}*{rc}{r}"
        r = tr_row[CTA]
        tr[f"{uc}{r}"] = f"=-SUM({uc}6:{uc}{r - 1})"
        for code, *_ in rows:
            for col, fmt, color in ((lc, NUM, "008000"), (rc, RATE, "008000"), (uc, NUM, "000000")):
                cell = tr[f"{col}{tr_row[code]}"]
                if cell.value is not None: cell.number_format = fmt; cell.font = _f(color=color)
        tr[f"{uc}{r}"].font = _f(True); tr[f"{uc}{r}"].border = TOP
    tr.cell(row=tr_row[CTA] + 2, column=2, value="Assets and liabilities at closing rate, income and expenses at average "
            "rate, share capital at historical rate, opening retained earnings at the USD amount carried forward. "
            "The translation reserve balances the translated trial balance.").font = _f(italic=True, color="595959")

    # Eliminations journal
    el = _sheet(wb, "Eliminations", "Consolidation adjustments (debit +, credit -)")
    _header(el, 5, ["#", "Description", "Code", "Account", "Amount"], widths=[5, 62, 8, 28, 12])
    name_of = {a[0]: a[1] for a in [(x.account_code, x.account_name) for x in coa.itertuples(index=False)] + [(g[0], g[1]) for g in GROUP_ACCOUNTS]}
    r = 6; entry_no = {}
    ent_by = entities.set_index("entity")
    sub_iter = list(entities.iloc[1:].itertuples(index=False))

    def usd(code, ent):
        return f"Translation!{tr_cols[ent]}{tr_row[code]}"
    formulas = []
    for e in sub_iter:
        er = ent_row[e.entity]; own, inv, pre = f"Inputs!$E${er}", f"Inputs!$F${er}", f"Inputs!$G${er}"
        n = f"Investment in {e.name}"
        sc = f"(-{usd(SC, e.entity)})"
        formulas += [(n, SC, f"={sc}"), (n, RE, f"={pre}"), (n, GW, f"={inv}-{own}*({sc}+{pre})"),
                     (n, INV, f"=-{inv}"), (n, NCI, f"=-(1-{own})*({sc}+{pre})")]
        if e.ownership < 1:
            n = f"NCI share of {e.name} post-acquisition reserves"
            post = f"(-{usd(RE, e.entity)}-{pre})"; cta = usd(CTA, e.entity)
            formulas += [(n, RE, f"=(1-{own})*{post}"), (n, NCI, f"=-(1-{own})*{post}"),
                         (n, CTA, f"=-(1-{own})*{cta}"), (n, NCI, f"=(1-{own})*{cta}")]
    for x, ir in zip(ic.itertuples(index=False), ic_row):
        amt = f"Inputs!$F${ir}"
        if x.type == "balance":
            n = f"Intercompany balance: {x.description}"; formulas += [(n, IC_PAY, f"={amt}"), (n, IC_REC, f"=-{amt}")]
        elif x.type == "sale":
            n = f"Intercompany sale: {x.description}"; formulas += [(n, IC_REV, f"={amt}"), (n, IC_PUR, f"=-{amt}")]
        elif x.type == "unrealised_profit":
            n = f"Unrealised profit in inventory: {x.description}"; formulas += [(n, COS, f"={amt}"), (n, INVENTORY, f"=-{amt}")]
    for n, code, f in formulas:
        entry_no.setdefault(n, len(entry_no) + 1)
        for j, v in enumerate([entry_no[n], n, code, name_of[code], f]):
            c = el.cell(row=r, column=2 + j, value=v); c.font = _f(color="000000" if j < 4 else "008000")
            if j == 4: c.number_format = NUM
        r += 1
    el_last = r - 1
    el.cell(row=r, column=5, value="Total (should be zero)").font = _f(True)
    c = el.cell(row=r, column=6, value=f"=ROUND(SUM(F6:F{el_last}),2)"); c.number_format = NUM; c.font = _f(True); c.border = TOP

    # Consolidation worksheet
    cs = _sheet(wb, "Consolidation", "Consolidation worksheet (debit +, credit -)")
    _header(cs, 5, ["Code", "Account"] + [f"{e} (USD)" for e in ents] + ["Aggregate", "Eliminations", "Consolidated"],
            widths=[8, 30] + [13] * (len(ents) + 3))
    cs_row = {}
    all_rows = [(a.account_code, a.account_name) for a in coa.itertuples(index=False)] + [(g[0], g[1]) for g in GROUP_ACCOUNTS]
    all_rows.sort()
    ecol = [cs.cell(row=5, column=4 + k).column_letter for k in range(len(ents))]
    agg, elc, con = (cs.cell(row=5, column=4 + len(ents) + m).column_letter for m in range(3))
    for i, (code, name) in enumerate(all_rows):
        r = 6 + i; cs_row[code] = r
        cs.cell(row=r, column=2, value=code).font = _f(); cs.cell(row=r, column=3, value=name).font = _f()
        for k, e in enumerate(ents):
            cs[f"{ecol[k]}{r}"] = f"=Translation!{tr_cols[e]}{tr_row[code]}" if code in tr_row else 0
        cs[f"{agg}{r}"] = f"=SUM({ecol[0]}{r}:{ecol[-1]}{r})"
        cs[f"{elc}{r}"] = f"=SUMIFS(Eliminations!$F$6:$F${el_last},Eliminations!$D$6:$D${el_last},$B{r})"
        cs[f"{con}{r}"] = f"={agg}{r}+{elc}{r}"
        for col in ecol + [agg, elc, con]:
            c = cs[f"{col}{r}"]; c.number_format = NUM
            c.font = _f(color="008000" if col in ecol or col == elc else "000000", bold=col == con)
    cs_last = 6 + len(all_rows) - 1
    cs.cell(row=cs_last + 1, column=3, value="Check (should be zero)").font = _f(True)
    for col in ecol + [agg, elc, con]:
        c = cs[f"{col}{cs_last + 1}"]; c.value = f"=ROUND(SUM({col}6:{col}{cs_last}),2)"; c.number_format = NUM
        c.font = _f(True); c.border = TOP

    # Presentation statements
    C = lambda code: f"Consolidation!{con}{cs_row[code]}"
    def T(code, ent): return f"Translation!{tr_cols[ent]}{tr_row[code]}"
    is_ = _sheet(wb, "Consolidated IS", "Consolidated income statement for the year ended 31 December 2025")
    is_.column_dimensions["B"].width = 44; is_.column_dimensions["C"].width = 14
    nci_terms = "+".join(f"(1-Inputs!$E${ent_row[e.entity]})*-SUM(Translation!{tr_cols[e.entity]}{tr_row[4000]}:"
                         f"{tr_cols[e.entity]}{tr_row[8000]})" for e in sub_iter)
    is_lines = [
        ("Revenue", f"=-({C(4000)}+{C(4050)})", None),
        ("Cost of sales", f"=-({C(5000)}+{C(5050)})", None),
        ("Gross profit", "=C6+C7", TOP),
        ("Operating expenses", f"=-{C(6000)}", None),
        ("Operating profit", "=C8+C9", TOP),
        ("Interest expense", f"=-{C(7000)}", None),
        ("Profit before tax", "=C10+C11", TOP),
        ("Income tax expense", f"=-{C(8000)}", None),
        ("Profit for the year", "=C12+C13", DBL),
        (None, None, None),
        ("Attributable to owners of the parent", "=C14-C17", None),
        ("Attributable to non-controlling interests", f"={nci_terms}", None),
    ]
    _header(is_, 5, ["", "2025"], widths=None)
    for i, (lab, f, border) in enumerate(is_lines):
        r = 6 + i
        if lab is None: continue
        is_[f"B{r}"] = lab; is_[f"C{r}"] = f; is_[f"C{r}"].number_format = NUM
        bold = border is not None
        is_[f"B{r}"].font = _f(bold); is_[f"C{r}"].font = _f(bold)
        if border: is_[f"C{r}"].border = border
    is_["B19"] = ("Unrealised profit is eliminated in full against the seller (Sub One, 100% owned), so it has no NCI effect.")
    is_["B19"].font = _f(italic=True, color="595959")

    bs = _sheet(wb, "Consolidated BS", "Consolidated balance sheet at 31 December 2025")
    bs.column_dimensions["B"].width = 44; bs.column_dimensions["C"].width = 14
    _header(bs, 5, ["", "2025"])
    bs_lines = [
        ("Goodwill", f"={C(GW)}", None), ("Property, plant and equipment", f"={C(1500)}", None),
        ("Total non-current assets", "=C6+C7", TOP),
        ("Inventory", f"={C(1200)}", None), ("Trade receivables", f"={C(1100)}", None),
        ("Intercompany receivable (eliminated)", f"={C(IC_REC)}", None), ("Cash", f"={C(1000)}", None),
        ("Total current assets", "=SUM(C9:C12)", TOP),
        ("Total assets", "=C8+C13", DBL),
        (None, None, None),
        ("Share capital", f"=-{C(SC)}", None),
        ("Retained earnings", f"=-{C(RE)}+'Consolidated IS'!C16", None),
        ("Translation reserve", f"=-{C(CTA)}", None),
        ("Equity attributable to owners of the parent", "=SUM(C16:C18)", TOP),
        ("Non-controlling interests", f"=-{C(NCI)}+'Consolidated IS'!C17", None),
        ("Total equity", "=C19+C20", TOP),
        ("Borrowings", f"=-{C(2500)}", None), ("Trade payables", f"=-{C(2000)}", None),
        ("Intercompany payable (eliminated)", f"=-{C(IC_PAY)}", None),
        ("Total liabilities", "=SUM(C22:C24)", TOP),
        ("Total equity and liabilities", "=C21+C25", DBL),
        (None, None, None),
        ("Check: assets less equity and liabilities", "=ROUND(C14-C26,2)", None),
    ]
    for i, (lab, f, border) in enumerate(bs_lines):
        r = 6 + i
        if lab is None: continue
        bs[f"B{r}"] = lab; bs[f"C{r}"] = f; bs[f"C{r}"].number_format = NUM
        bold = border is not None
        bs[f"B{r}"].font = _f(bold); bs[f"C{r}"].font = _f(bold)
        if border: bs[f"C{r}"].border = border

    # Checks
    ck = _sheet(wb, "Checks", "Checks (all should be zero)")
    ck.column_dimensions["B"].width = 52; ck.column_dimensions["C"].width = 14
    checks = [("Entity trial balances balance",
               "+".join(f"ABS('Trial Balances'!{tbs.cell(row=5, column=5 + k).column_letter}{last_tb + 1})" for k in range(len(ents)))),
              ("Elimination journal balances", f"ABS(Eliminations!F{el_last + 1})"),
              ("Consolidation worksheet balances", f"ABS(Consolidation!{con}{cs_last + 1})"),
              ("Intercompany receivables and payables fully eliminated", f"ABS({C(IC_REC)})+ABS({C(IC_PAY)})"),
              ("Intercompany sales and purchases fully eliminated", f"ABS({C(IC_REV)})+ABS({C(IC_PUR)})"),
              ("Balance sheet balances", "ABS('Consolidated BS'!C28)")]
    for i, (lab, f) in enumerate(checks):
        ck[f"B{6 + i}"] = lab; ck[f"B{6 + i}"].font = _f()
        ck[f"C{6 + i}"] = f"=ROUND({f},2)"; ck[f"C{6 + i}"].number_format = NUM; ck[f"C{6 + i}"].font = _f()
    r = 6 + len(checks) + 1
    ck[f"B{r}"] = "All checks pass?"; ck[f"B{r}"].font = _f(True)
    ck[f"C{r}"] = f'=IF(SUM(C6:C{r - 2})=0,"OK","ERROR")'; ck[f"C{r}"].font = _f(True)

    order = ["Consolidated IS", "Consolidated BS", "Consolidation", "Eliminations", "Translation", "Trial Balances", "Inputs", "Checks"]
    wb._sheets = [wb[n] for n in order]
    for ws in wb.worksheets:
        ws.page_setup.orientation = "landscape"; ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    wb.save(out_path)


# ---------------------------------------------------------------- main
def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    p.add_argument("--data", default="data")
    p.add_argument("--out", default="output/group_consolidation.xlsx")
    a = p.parse_args(argv)
    try:
        wide, coa, entities, fx, ic = load_inputs(a.data)
        usd, journal, elim, consolidated = consolidate(wide, coa, entities, fx, ic)
        s = summarise(usd, entities, consolidated)
        if abs(s["Total assets"] - s["Total equity and liabilities"]) > 0.01:
            raise ValidationError("Consolidated balance sheet does not balance")
    except ValidationError as e:
        print(f"Stopped: {e}"); return 1
    Path(a.out).parent.mkdir(parents=True, exist_ok=True)
    write_workbook(wide, coa, entities, fx, ic, journal, a.out)
    print(f"Checks passed: {len(entities)} trial balances balance, {len({j[0] for j in journal})} elimination entries "
          f"balance, intercompany fully eliminated, balance sheet balances.")
    for k, v in s.items():
        print(f"  {k:<32}{v:>10,.1f}")
    print(f"Saved {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
