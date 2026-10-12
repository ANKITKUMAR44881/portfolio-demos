"""
E&O engine - pure pandas/openpyxl, no UI. Importable + unit-testable.
Inputs: CTB workbook (Inventory Table) + MRP workbook (MRP Purchasing 12Weeks).
"""
import io
import re
import os
import glob
import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, LineChart, Reference

_MON = {m: i + 1 for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"])}
_MON_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def parse_month(h):
    """Return (year, month) from a header cell that is a date or a string like 'Aug-26' / 'AUG 2026'."""
    if isinstance(h, (datetime.datetime, datetime.date)):
        return (h.year, h.month)
    s = str(h).strip().lower()
    if not s or s == "none":
        return None
    mo = next((v for k, v in _MON.items() if k in s), None)
    if not mo:
        return None
    y4 = re.search(r"(20\d{2})", s)
    if y4:
        return (int(y4.group(1)), mo)
    y2 = re.search(r"[-/ ](\d{2})\b", s)
    if y2:
        return (2000 + int(y2.group(1)), mo)
    return None


def _ym_minus(ym, n):
    y, m = ym
    m -= n
    while m <= 0:
        m += 12
        y -= 1
    return (y, m)


def _ym_label(ym):
    return f"{_MON_ABBR[ym[1] - 1]}-{str(ym[0])[2:]}"


def _cycle_label(fname):
    base = os.path.basename(fname).replace(".xlsx", "")
    months = re.findall(r"([A-Za-z]{3,9})\s*20(\d{2})", base)
    if len(months) >= 2:
        return f"{months[0][0][:3].title()}{months[0][1]}-{months[1][0][:3].title()}{months[1][1]}"
    return base[:18]


def clean_family(v):
    """'01 PRODUCT NAME' -> 'Product Name'; blanks/None -> 'Unmapped'."""
    s = str(v).strip()
    if s == "" or s.lower() in ("none", "nan", "#n/a"):
        return "Unmapped"
    s = re.sub(r"^\d+\s+", "", s)            # drop leading sort-number prefix
    return s.title()

BLUE = "F2E1D9"; GREEN = "C6E7CE"; ORANGE = "F79D46"; YELLOW = "FFFF99"
HDR = Font(bold=True)


def _src(file_or_bytes):
    return io.BytesIO(file_or_bytes) if isinstance(file_or_bytes, (bytes, bytearray)) else file_or_bytes


def _find_sheet(xls, *needles):
    for s in xls.sheet_names:
        low = s.lower()
        if all(n in low for n in needles):
            return s
    return None


def load_inventory(file_or_bytes):
    xls = pd.ExcelFile(_src(file_or_bytes), engine="openpyxl")
    inv_sheet = _find_sheet(xls, "inventory", "table") or _find_sheet(xls, "inventory")
    df = xls.parse(inv_sheet)
    df.columns = [str(c).strip() for c in df.columns]

    def col(*names):
        for n in names:
            for c in df.columns:
                if c.lower() == n.lower():
                    return c
        return None

    c_pn, c_desc = col("PART_NO"), col("PART_DESC")
    c_qty, c_uv = col("QTY_ONHAND"), col("INV_VALUE")
    c_ext, c_fam = col("EXT_COST"), col("Family")
    c_type, c_last = col("Part Type"), col("LAST_ACTIVITY")
    c_rec = col("REC_DATE")

    out = pd.DataFrame({
        "PART_NO": df[c_pn].astype(str).str.strip(),
        "PART_DESC": df[c_desc] if c_desc else "",
        "QTY": pd.to_numeric(df[c_qty], errors="coerce").fillna(0),
        "UNIT_VAL": pd.to_numeric(df[c_uv], errors="coerce").fillna(0) if c_uv else 0,
        "EXT": pd.to_numeric(df[c_ext], errors="coerce") if c_ext else None,
        "FAMILY": df[c_fam] if c_fam else "Unmapped",
        "PTYPE": df[c_type].astype(str) if c_type else "#N/A",
        "LAST_ACT": df[c_last] if c_last else "",
        "REC_DATE": pd.to_datetime(df[c_rec], errors="coerce") if c_rec else pd.NaT,
    })
    if c_ext is None:
        out["EXT"] = out["UNIT_VAL"] * out["QTY"]
    else:
        out["EXT"] = out["EXT"].fillna(out["UNIT_VAL"] * out["QTY"])

    pn = out["PART_NO"].str.lower()
    suffix_bad = pn.str.endswith("-f") | pn.str.endswith("-hb") | pn.str.endswith("-hw")
    t = out["PTYPE"].str.strip().str.lower()
    type_ok = t.eq("company owned")   # company-owned ONLY: excludes consigned, harvested, AND blank/#N/A
    elig = out[(~suffix_bad) & type_ok].copy()

    agg = (elig.groupby("PART_NO", as_index=False)
           .agg(PART_DESC=("PART_DESC", "first"),
                IOH_QTY=("QTY", "sum"),
                IOH_DOLLARS=("EXT", "sum"),
                UNIT_VAL=("UNIT_VAL", "max"),
                FAMILY=("FAMILY", "first"),
                LAST_ACT=("LAST_ACT", "max"),
                FIRST_REC=("REC_DATE", "min")))

    # authoritative product family from the Component Mapping sheet (PART_NO -> Family)
    cm_sheet = _find_sheet(xls, "component", "mapping")
    if cm_sheet:
        cm = xls.parse(cm_sheet)
        cm.columns = [str(c).strip() for c in cm.columns]
        pn_c = next((c for c in cm.columns if c.lower() in ("part_no", "part no", "partno")), cm.columns[0])
        fam_c = next((c for c in cm.columns if c.lower() == "family"), None)
        if fam_c:
            fam_map = dict(zip(cm[pn_c].astype(str).str.strip(), cm[fam_c]))
            mapped = agg["PART_NO"].map(fam_map)
            agg["FAMILY"] = mapped.where(mapped.notna(), agg["FAMILY"])
    agg["FAMILY"] = agg["FAMILY"].map(clean_family)
    return agg, {"rows_total": len(out), "rows_eligible": len(elig), "parts": len(agg), "sheet": inv_sheet}


# ---- Consumables / MRO (NOT E&O; billed to the customer via monthly chargebacks) ----
# Supply part-number pattern treated as consumables, not service inventory (demo pattern).
CONS_RE = re.compile(r"^S-\d", re.I)   # demo: consumable/supply PNs start with "S-"
# Name match kept deliberately narrow - broad words would also match real BOM components.
CONS_NAME_RE = re.compile(r"consumable supply", re.I)


def load_consumable_pns(chargeback_paths=None, master_path=None):
    """Union of part numbers billed as consumables in the chargeback files
    (rolling window the caller passes) plus an optional consumables master list."""
    pns = set()
    for f in (chargeback_paths or []):
        try:
            df = pd.read_excel(f, sheet_name=0, dtype=str)
        except Exception:
            continue
        df.columns = [str(c).strip() for c in df.columns]
        c = next((x for x in df.columns if x.lower() in ("part no", "part_no", "partno", "part number")), None)
        if not c:
            continue
        for v in df[c].dropna().astype(str):
            v = v.strip()
            if v and v.lower() != "nan":
                pns.add(v)
    if master_path:
        try:
            mx = pd.ExcelFile(master_path)
            for s in mx.sheet_names:
                d = mx.parse(s, dtype=str)
                d.columns = [str(x).strip() for x in d.columns]
                c = next((x for x in d.columns if x.lower() in ("part no", "part_no", "partno", "master pn", "pn")), None)
                if c:
                    for v in d[c].dropna().astype(str):
                        v = v.strip()
                        if v and v.lower() != "nan":
                            pns.add(v)
        except Exception:
            pass
    return pns


def is_consumable(part_no, desc="", extra=None):
    pn = str(part_no).strip()
    if extra and pn in extra:
        return True
    if CONS_RE.match(pn):
        return True
    if CONS_NAME_RE.search(str(desc)):
        return True
    return False


def split_eligible(agg, consumables=None, new_product_pns=None, run_date=None, seed_days=90):
    """Carve the aggregated owned inventory into three buckets:
       - E&O-eligible (returned first),
       - Consumables/MRO (billed via chargeback; never E&O),
       - Seed stock / new product (new program, or first received < seed_days ago; not E&O yet).
    Returns (eligible_df, consumables_df, seed_df)."""
    consumables = consumables or set()
    new_product_pns = set(new_product_pns or set())
    a = agg.copy()

    cons_mask = a.apply(lambda r: is_consumable(r["PART_NO"], r.get("PART_DESC", ""), consumables), axis=1)

    # Seed stock = strictly parts on the new-product list. The 90-day window is
    # the GRADUATION clock: a listed new-product part leaves seed once it is older than seed_days.
    # There is NO blanket "recently received" gate - that mis-flags normal replenishment.
    seed_mask = a["PART_NO"].isin(new_product_pns)
    if run_date is not None and "FIRST_REC" in a.columns and seed_mask.any():
        rd = pd.Timestamp(run_date)
        age_days = (rd - pd.to_datetime(a["FIRST_REC"], errors="coerce")).dt.days
        graduated = seed_mask & age_days.notna() & (age_days >= seed_days)
        seed_mask = seed_mask & ~graduated
    seed_mask = seed_mask & (~cons_mask)   # consumable classification wins

    cons_df = a[cons_mask].copy()
    cons_df["REASON"] = "Consumable / MRO - billed via customer chargeback, excluded from E&O"
    seed_df = a[seed_mask].copy()
    seed_df["REASON"] = f"Seed stock / new product - not E&O until {seed_days} days after first receipt"
    eligible = a[~(cons_mask | seed_mask)].copy()
    return eligible, cons_df, seed_df


def load_demand(file_or_bytes):
    xls = pd.ExcelFile(_src(file_or_bytes), engine="openpyxl")
    sht = _find_sheet(xls, "mrp purchasing 12") or _find_sheet(xls, "purchasing", "12")
    df = xls.parse(sht, header=1)
    df.columns = [str(c).strip() for c in df.columns]

    def col(*names):
        for n in names:
            for c in df.columns:
                if c.lower() == n.lower():
                    return c
        return None

    c_pn, c_dem, c_fam = col("Component"), col("12 Week Demand"), col("Product Family")
    c_own = col("Company Owned")   # MRP owned on-hand qty (authoritative owned split)
    d = pd.DataFrame({
        "PART_NO": df[c_pn].astype(str).str.strip(),
        "DEM12": pd.to_numeric(df[c_dem], errors="coerce").fillna(0),
        "PROD_FAM": df[c_fam] if c_fam else "",
        "MRP_OWN_QTY": pd.to_numeric(df[c_own], errors="coerce").fillna(0) if c_own else 0,
    })
    d = d[d["PART_NO"].str.len() > 0]
    return d.groupby("PART_NO", as_index=False).agg(
        DEM12=("DEM12", "sum"), PROD_FAM=("PROD_FAM", "first"),
        MRP_OWN_QTY=("MRP_OWN_QTY", "max")), sht


def run_eo(inv, dem, markup=0.05, materiality=50.0):
    df = inv.merge(dem, on="PART_NO", how="left")
    df["IN_MRP"] = df["DEM12"].notna()
    df["DEM12"] = df["DEM12"].fillna(0)
    # cross-check CTB company-owned IOH against MRP owned on-hand
    if "MRP_OWN_QTY" in df.columns:
        df["MRP_OWN_QTY"] = df["MRP_OWN_QTY"].fillna(0)
        df["OWN_RECON_DIFF"] = df["IOH_QTY"] - df["MRP_OWN_QTY"]
    else:
        df["MRP_OWN_QTY"] = 0.0
        df["OWN_RECON_DIFF"] = df["IOH_QTY"]
    df["DEM13"] = (df["DEM12"] * 13.0 / 12.0).round(0)
    df["EOH_QTY"] = (df["IOH_QTY"] - df["DEM13"]).clip(lower=0)
    df["EOH_DOL"] = df["EOH_QTY"] * df["UNIT_VAL"]
    df["EOH_BILL"] = df["EOH_DOL"] * (1 + markup)
    df["PRODUCT"] = df["FAMILY"].where(df["FAMILY"].astype(str).ne("Unmapped"), df["PROD_FAM"])
    df["PRODUCT"] = df["PRODUCT"].map(clean_family)

    def disposition(r):
        if r["DEM13"] > 0:
            return "PLEASE APPROVE CONSIGNMENT"
        if not r["IN_MRP"]:
            return "REVIEW - demand-linkage unverified"
        return "PLEASE APPROVE CONSIGNMENT"

    def confidence(r):
        return "Low" if (not r["IN_MRP"] or r["DEM13"] == 0) else "Medium"

    def reason(r):
        if not r["IN_MRP"]:
            return ("Part has on-hand but no demand row in MRP - verify parent BOM completeness and "
                    "demand-plan presence before reporting as E&O (possible incomplete BOM).")
        if r["DEM13"] == 0:
            return (f"No demand next 13 weeks; {int(r['EOH_QTY'])} units on hand. "
                    "Verify EOL/SBOM status - scrap if obsolete, else consign.")
        return (f"Demand next 13 weeks = {int(r['DEM13'])} units vs IOH {int(r['IOH_QTY'])}; "
                f"{int(r['EOH_QTY'])} units excess.")

    def cause(r):
        if not r["IN_MRP"]:
            return "Unverified (BOM linkage)"
        if r["DEM13"] == 0:
            return "No demand"
        return "Demand drop"

    eo = df[df["EOH_QTY"] > 0].copy()
    if len(eo):
        eo["REQUEST"] = eo.apply(disposition, axis=1)
        eo["CONFIDENCE"] = eo.apply(confidence, axis=1)
        eo["REASON"] = eo.apply(reason, axis=1)
        eo["CAUSE"] = eo.apply(cause, axis=1)
    else:
        for c in ["REQUEST", "CONFIDENCE", "REASON", "CAUSE"]:
            eo[c] = []
    eo["EVIDENCE"] = "CTB Inventory Table (QTY_ONHAND, INV_VALUE); MRP 12Wk Demand col O x 13/12"
    eo = eo[eo["EOH_BILL"] >= materiality].sort_values("EOH_BILL", ascending=False)
    return eo


def kpis(eo):
    if len(eo) == 0:
        return {"parts": 0, "eoh_dollars": 0.0, "billable": 0.0, "unverified": 0, "consign": 0.0}
    return {
        "parts": int(len(eo)),
        "eoh_dollars": float(eo["EOH_DOL"].sum()),
        "billable": float(eo["EOH_BILL"].sum()),
        "unverified": int(eo["REQUEST"].str.contains("unverified").sum()),
        "consign": float(eo.loc[eo["REQUEST"].str.contains("CONSIGN"), "EOH_BILL"].sum()),
    }


LEGEND = "Color legend: CTB inventory = blue | MRP demand = orange | Calculated = green | Agent/Manual = yellow"

FIELD_GUIDE = [
    ("A", "Product", "Mapping", "Part mapped to product/program via CTB Component Mapping (PART_NO -> Family)."),
    ("B", "Customer PN", "CTB Inventory Table", "PART_NO. Eligible only if Part Type = Company Owned; exclude -f/-hb/-hw."),
    ("C", "Customer PN Description", "CTB Inventory Table", "PART_DESC."),
    ("D", "IOH (Qty)", "CTB Inventory Table", "Sum of QTY_ONHAND by part."),
    ("E", "IOH ($)", "Calculated", "Sum of EXT_COST (INV_VALUE x QTY_ONHAND)."),
    ("F", "Demand - Next 13 Weeks", "MRP", "MRP Purchasing 12Weeks '12 Week Demand' x 13/12."),
    ("G", "EOH After 13 Weeks (Qty)", "Calculated", "MAX(0, IOH Qty - 13-week demand)."),
    ("H", "EOH After 13 Weeks ($)", "Calculated", "EOH Qty x Unit Inv Value."),
    ("I", "EOH After 13 Weeks ($) w/5%", "Calculated", "EOH $ x 1.05. Billable to the customer."),
    ("J", "Request to Consign/Scrap/Transfer", "Agent (human-approved)", "v1 defaults to Consign; verify EOL for scrap."),
    ("K", "Reason for E&O", "Agent (human-approved)", "Short customer-ready narrative + evidence."),
    ("L", "Confidence", "Agent", "High/Med/Low; Low routes to planner review."),
    ("M", "Evidence / Source", "Agent", "Tabs/columns used to derive the line."),
    ("N", "Customer Feedback", "Manual", "Filled after the customer responds."),
    ("O", "Unit Inv Value", "CTB Inventory Table", "INV_VALUE."),
    ("P", "Last Purchase Month", "ERP open PO report", "v2 - most recent purchase month."),
    ("Q", "Inbound On-Order Qty", "Open POs report", "v2 - firm on-order qty added to IOH."),
    ("R", "Aging / Last Activity", "CTB Inventory Table", "LAST_ACTIVITY date; drives aged/no-usage."),
]

METHODOLOGY = [
    "Horizon = 13 weeks (MRP 12-week demand x 13/12). Billing markup = 5%.",
    "Eligibility = Company Owned only; consigned (-f) and harvested (-hb/-hw) excluded.",
    "Product family from CTB Component Mapping (PART_NO -> Family).",
    "Coverage guardrail = zero-demand parts with no MRP row flagged 'unverified' (possible incomplete BOM), not auto-reported.",
    "v1 scope = deterministic calc + KPIs. v2 = EOL scrap logic, demand waterfall, open POs, alternates, inter-site transfer.",
    "Sources & as-of = CTB inventory refresh; MRP demand-plan version.",
]

HEADERS = ["Product", "Customer PN", "Customer PN Description", "IOH (Qty)", "IOH ($)",
           "Demand - Next 13 Weeks (Qty)", "EOH After 13 Weeks (Qty)", "EOH After 13 Weeks ($)",
           "EOH After 13 Weeks ($) w/ 5% Markup", "Request to Consign/Scrap/Transfer", "Reason for E&O",
           "Confidence", "Evidence / Source", "Customer Feedback", "Unit Inv Value",
           "Last Purchase Month", "Inbound On-Order Qty", "Aging / Last Activity"]
HEADER_FILLS = [BLUE, BLUE, BLUE, BLUE, GREEN, ORANGE, GREEN, GREEN, GREEN, YELLOW, YELLOW, YELLOW,
                YELLOW, YELLOW, BLUE, BLUE, ORANGE, BLUE]


def _hdr_cell(c, fill):
    c.font = HDR; c.fill = PatternFill("solid", fgColor=fill); c.alignment = Alignment(wrap_text=True)


def _working_tab(wb, eo, site):
    ws = wb.active
    ws.title = f"Quarterly E&O Update - {site}"[:31]
    ws.cell(1, 1, f"Quarterly E&O Service Materials - {site}    |    {LEGEND}").font = HDR
    for j, name in enumerate(HEADERS, start=1):
        _hdr_cell(ws.cell(2, j, name), HEADER_FILLS[j - 1])
    for i, (_, r) in enumerate(eo.iterrows(), start=3):
        ws.cell(i, 1, r["PRODUCT"]); ws.cell(i, 2, r["PART_NO"]); ws.cell(i, 3, r["PART_DESC"])
        ws.cell(i, 4, round(float(r["IOH_QTY"]), 0))
        ws.cell(i, 5, round(float(r["IOH_DOLLARS"]), 2)).number_format = '$#,##0.00'
        ws.cell(i, 6, int(r["DEM13"])); ws.cell(i, 7, int(r["EOH_QTY"]))
        ws.cell(i, 8, round(float(r["EOH_DOL"]), 2)).number_format = '$#,##0.00'
        ws.cell(i, 9, round(float(r["EOH_BILL"]), 2)).number_format = '$#,##0.00'
        ws.cell(i, 10, r["REQUEST"]); ws.cell(i, 11, r["REASON"]); ws.cell(i, 12, r["CONFIDENCE"])
        ws.cell(i, 13, r["EVIDENCE"]); ws.cell(i, 15, round(float(r["UNIT_VAL"]), 4))
        ws.cell(i, 18, str(r["LAST_ACT"]))
    for ci, w in {1: 22, 2: 18, 3: 40, 10: 30, 11: 55, 13: 45}.items():
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = "A3"


def _mini_table(s, top, title, pairs, money=True):
    s.cell(top, 1, title).font = HDR
    s.cell(top + 1, 1, "Category"); s.cell(top + 1, 2, "Billable $")
    for i, (name, val) in enumerate(pairs):
        s.cell(top + 2 + i, 1, name)
        c = s.cell(top + 2 + i, 2, round(float(val), 2))
        if money:
            c.number_format = '$#,##0'
    return top + 1 + len(pairs)  # last data row


def _summary_tab(wb, eo, site, run_date):
    s = wb.create_sheet(f"Summary - {site}"[:31])
    k = kpis(eo)
    s["A1"] = f"E&O Summary - {site}"; s["A1"].font = Font(bold=True, size=14)
    s["A2"] = f"Run date {run_date}  |  Horizon 13 weeks  |  Markup 5%"
    # KPI tiles
    kpi = [("E&O parts", k["parts"]), ("Total EOH $", k["eoh_dollars"]), ("Billable $ (w/5%)", k["billable"]),
           ("Consign $ (w/5%)", k["consign"]), ("Demand-linkage UNVERIFIED", k["unverified"])]
    for j, (lbl, val) in enumerate(kpi, start=1):
        _hdr_cell(s.cell(4, j, lbl), BLUE)
        c = s.cell(5, j, round(val, 0) if isinstance(val, float) else val)
        if isinstance(val, float):
            c.number_format = '$#,##0'
    # methodology
    s.cell(7, 1, "Methodology & Assumptions").font = HDR
    for i, m in enumerate(METHODOLOGY):
        s.cell(8 + i, 1, "- " + m)

    if len(eo):
        prod = list(eo.groupby("PRODUCT")["EOH_BILL"].sum().sort_values(ascending=False).head(10).items())
        cause = list(eo.groupby("CAUSE")["EOH_BILL"].sum().sort_values(ascending=False).items())
        disp = list(eo.groupby("REQUEST")["EOH_BILL"].sum().sort_values(ascending=False).items())
        top = (eo.sort_values("EOH_BILL", ascending=False)
               .head(10)[["PART_NO", "PRODUCT", "CAUSE", "EOH_BILL"]].values.tolist())

        r1 = 16
        last1 = _mini_table(s, r1, "EOH $ by Product (top 10)", prod)
        bc = BarChart(); bc.type = "bar"; bc.title = "EOH $ by Product (top 10)"; bc.legend = None
        bc.add_data(Reference(s, min_col=2, min_row=r1 + 1, max_row=last1), titles_from_data=True)
        bc.set_categories(Reference(s, min_col=1, min_row=r1 + 2, max_row=last1))
        bc.height = 8; bc.width = 16; s.add_chart(bc, "D" + str(r1))

        r2 = last1 + 3
        last2 = _mini_table(s, r2, "EOH $ by Cause", cause)
        pie = PieChart(); pie.title = "EOH $ by Cause"
        pie.add_data(Reference(s, min_col=2, min_row=r2 + 1, max_row=last2), titles_from_data=True)
        pie.set_categories(Reference(s, min_col=1, min_row=r2 + 2, max_row=last2))
        pie.height = 8; pie.width = 16; s.add_chart(pie, "D" + str(r2))

        r3 = last2 + 3
        last3 = _mini_table(s, r3, "Billable $ by Disposition", disp)
        bc2 = BarChart(); bc2.type = "col"; bc2.title = "Billable $ by Disposition"; bc2.legend = None
        bc2.add_data(Reference(s, min_col=2, min_row=r3 + 1, max_row=last3), titles_from_data=True)
        bc2.set_categories(Reference(s, min_col=1, min_row=r3 + 2, max_row=last3))
        bc2.height = 8; bc2.width = 16; s.add_chart(bc2, "D" + str(r3))

        r4 = last3 + 3
        s.cell(r4, 1, "Top E&O $ drivers").font = HDR
        for j, h in enumerate(["Customer PN", "Product", "Cause", "Billable $"], start=1):
            s.cell(r4 + 1, j, h).font = HDR
        for i, row in enumerate(top):
            s.cell(r4 + 2 + i, 1, row[0]); s.cell(r4 + 2 + i, 2, row[1]); s.cell(r4 + 2 + i, 3, row[2])
            s.cell(r4 + 2 + i, 4, round(float(row[3]), 2)).number_format = '$#,##0'
    s.column_dimensions["A"].width = 28; s.column_dimensions["B"].width = 16; s.column_dimensions["C"].width = 22


def _field_guide(wb):
    fg = wb.create_sheet("Field Guide")
    fg.cell(1, 1, "Field Guide - source & derivation of each working-tab column").font = HDR
    for j, h in enumerate(["Column", "Field", "Source", "How it is derived"], start=1):
        _hdr_cell(fg.cell(3, j, h), BLUE)
    for i, row in enumerate(FIELD_GUIDE, start=4):
        for j, v in enumerate(row, start=1):
            fg.cell(i, j, v)
    for ci, w in {1: 8, 2: 32, 3: 22, 4: 70}.items():
        fg.column_dimensions[get_column_letter(ci)].width = w


def _global_summary(wb, eo, site):
    gs = wb.create_sheet("Global Summary")
    gs.cell(1, 1, "Global E&O Summary - Consign/Scrap/Transfer by site").font = HDR
    cols = ["Site", "Consign Qty", "Consign $ w/5%", "Scrap Qty", "Scrap $ w/5%",
            "Transfer Qty", "Transfer $ w/5%", "Total EOH $ w/5%"]
    for j, h in enumerate(cols, start=1):
        _hdr_cell(gs.cell(3, j, h), BLUE)

    def bucket(key):
        m = eo["REQUEST"].str.contains(key, case=False)
        return float(eo.loc[m, "EOH_QTY"].sum()), float(eo.loc[m, "EOH_BILL"].sum())
    cq, cd = bucket("CONSIGN"); sq, sd = bucket("SCRAP"); tq, td = bucket("TRANSFER")
    tot = float(eo["EOH_BILL"].sum()) if len(eo) else 0
    sites = ["SITE-1", "SITE-2", "SITE-3", "GLOBAL"]
    for i, st in enumerate(sites, start=4):
        gs.cell(i, 1, st)
        if st == site or st == "GLOBAL":
            vals = [cq, cd, sq, sd, tq, td, tot]
            for j, v in enumerate(vals, start=2):
                c = gs.cell(i, j, round(v, 2))
                if j in (3, 5, 7, 8):
                    c.number_format = '$#,##0'
    for ci in range(1, 9):
        gs.column_dimensions[get_column_letter(ci)].width = 16


def _waterfall_tab(wb):
    dw = wb.create_sheet("Demand Waterfall")
    dw.cell(1, 1, "Demand Waterfall - lead-time-aware drop detection (v2: not yet wired)").font = HDR
    dw.cell(2, 1, "v2 will populate Plan T-2 -> T-1 -> Current by product and chart the drop that explains the E&O.")
    for j, h in enumerate(["Product / OPN", "Lead time (wks)", "Plan T-2 (drove buy)", "Plan T-1",
                           "Current Plan", "Delta", "% Drop", "Likely driver"], start=1):
        _hdr_cell(dw.cell(4, j, h), BLUE)


def load_demand_snapshots(folder, site="SITE-1"):
    """Read every demand-plan workbook in `folder`; return cycles sorted oldest->newest.
    Each cycle = {label, mtime, model_col, month_cols:{col:(y,m)}, df(filtered to site warehouse)}."""
    snaps = []
    for path in sorted(glob.glob(os.path.join(folder, "*.xlsx"))):
        if os.path.basename(path).startswith("~$"):
            continue
        try:
            xls = pd.ExcelFile(path, engine="openpyxl")
        except Exception:
            continue
        sht = next((s for s in xls.sheet_names if s.strip().lower() == "demand plan"), None)
        if not sht:
            continue
        raw = xls.parse(sht, header=3)              # header is row 4
        cols = list(raw.columns)

        def findc(name):
            return next((c for c in cols if str(c).strip().lower() == name), None)
        wh, mdl = findc("warehouse"), findc("model")
        if wh is None or mdl is None:
            continue
        month_cols = {c: parse_month(c) for c in cols if parse_month(c)}
        if not month_cols:
            continue
        df = raw[raw[wh].astype(str).str.strip().str.upper() == site.upper()].copy()
        snaps.append({"label": _cycle_label(path), "mtime": os.path.getmtime(path),
                      "mdl": mdl, "months": month_cols, "df": df})
    snaps.sort(key=lambda x: x["mtime"])
    return snaps


def build_waterfall(snaps, offset_months=1, top_n=8):
    """Track demand for ONE target planning month across cycles (customer month - offset)."""
    if not snaps:
        return None
    target_cust = min(snaps[-1]["months"].values())          # earliest forward month of newest plan
    recon_label = _ym_label(_ym_minus(target_cust, offset_months))
    series = []
    for s in snaps:
        col = next((c for c, pm in s["months"].items() if pm == target_cust), None)
        if col is None:
            continue
        d = s["df"][[s["mdl"], col]].copy()
        d[col] = pd.to_numeric(d[col], errors="coerce").fillna(0)
        series.append((s["label"], d.groupby(s["mdl"])[col].sum()))
    if len(series) < 2:
        return None
    table = pd.DataFrame({lbl: g for lbl, g in series}).fillna(0)
    table = table.loc[table.sum(axis=1) > 0]
    last = series[-1][0]
    table = table.sort_values(last, ascending=False)
    top = table.head(top_n)
    if len(table) > top_n:
        other = table.iloc[top_n:].sum(); other.name = "Other"
        top = pd.concat([top, other.to_frame().T])
    return {"table": top, "target": recon_label, "cycles": [l for l, _ in series]}


def _waterfall_tab(wb, wf=None):
    dw = wb.create_sheet("Demand Waterfall")
    if not wf:
        dw.cell(1, 1, "Demand Waterfall - no demand-plan snapshots loaded").font = HDR
        dw.cell(2, 1, "Load demand-plan snapshots to populate this.")
        for j, h in enumerate(["Product / Model", "Plan T-2", "Plan T-1", "Current"], start=1):
            _hdr_cell(dw.cell(4, j, h), BLUE)
        return
    dw.cell(1, 1, f"Demand Waterfall - demand for {wf['target']} as forecast across plan cycles").font = HDR
    dw.cell(2, 1, "Customer month offset -1 applied (customer Aug = planning Jul). "
                  "Falling left-to-right = demand drop = E&O signal.")
    tbl = wf["table"]; cycles = list(tbl.columns)
    _hdr_cell(dw.cell(4, 1, "Product / Model"), BLUE)
    for j, cy in enumerate(cycles, start=2):
        _hdr_cell(dw.cell(4, j, cy), ORANGE)
    for i, (model, row) in enumerate(tbl.iterrows(), start=5):
        dw.cell(i, 1, str(model))
        for j, cy in enumerate(cycles, start=2):
            dw.cell(i, j, int(round(float(row[cy]))))
    last = 4 + len(tbl)
    lc = LineChart(); lc.title = f"Demand for {wf['target']} across plan cycles"
    lc.y_axis.title = "Units"; lc.x_axis.title = "Plan cycle"; lc.height = 9; lc.width = 20
    data = Reference(dw, min_col=1, max_col=1 + len(cycles), min_row=5, max_row=last)
    lc.add_data(data, titles_from_data=True, from_rows=True)
    lc.set_categories(Reference(dw, min_col=2, max_col=1 + len(cycles), min_row=4, max_row=4))
    dw.add_chart(lc, "A" + str(last + 3))
    dw.column_dimensions["A"].width = 24


def _carveout_tab(wb, title, df, note):
    ws = wb.create_sheet(title[:31])
    ws.cell(1, 1, note).font = HDR
    heads = ["Product / Family", "Customer PN", "Description", "IOH (Qty)", "IOH ($)", "Unit Inv Value",
             "First Received", "Reason"]
    for j, h in enumerate(heads, 1):
        _hdr_cell(ws.cell(2, j, h), YELLOW)
    tot = 0.0
    for i, (_, r) in enumerate(df.iterrows(), start=3):
        ws.cell(i, 1, str(r.get("FAMILY", "")))
        ws.cell(i, 2, r["PART_NO"]); ws.cell(i, 3, str(r.get("PART_DESC", "")))
        ws.cell(i, 4, round(float(r.get("IOH_QTY", 0)), 0))
        c = ws.cell(i, 5, round(float(r.get("IOH_DOLLARS", 0) or 0), 2)); c.number_format = '$#,##0.00'
        ws.cell(i, 6, round(float(r.get("UNIT_VAL", 0) or 0), 4))
        fr = r.get("FIRST_REC", None)
        ws.cell(i, 7, str(fr)[:10] if fr is not None and str(fr) != "NaT" else "")
        ws.cell(i, 8, str(r.get("REASON", "")))
        tot += float(r.get("IOH_DOLLARS", 0) or 0)
    ws.cell(2, 5)  # header already set
    n = len(df)
    ws.cell(3 + n if n else 3, 4, "TOTAL").font = HDR
    tc = ws.cell(3 + n if n else 3, 5, round(tot, 2)); tc.number_format = '$#,##0.00'; tc.font = HDR
    for ci, w in {1: 20, 2: 16, 3: 42, 8: 55}.items():
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = "A3"


def build_workbook(eo, site, run_date, wf=None, consumables=None, seed=None):
    wb = openpyxl.Workbook()
    _working_tab(wb, eo, site)
    _summary_tab(wb, eo, site, run_date)
    if consumables is not None and len(consumables):
        _carveout_tab(wb, f"Consumables - {site}", consumables,
                      "Consumables / MRO - billed to the customer via monthly chargeback. EXCLUDED from E&O math.")
    if seed is not None and len(seed):
        _carveout_tab(wb, f"Seed Stock - {site}", seed,
                      "Seed stock / new product - not E&O until after the first 90-day cycle. EXCLUDED from E&O math.")
    _field_guide(wb)
    _global_summary(wb, eo, site)
    _waterfall_tab(wb, wf)
    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
    return bio
