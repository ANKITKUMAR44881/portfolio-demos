"""
Synthetic sample data for the E&O demo. Everything here is randomly generated -
no real part numbers, products, customers, sites, or prices.
"""
import io
import datetime as dt
import numpy as np
import pandas as pd

PRODUCTS = ["Headset Alpha", "Headset Beta", "Controller Pro", "Smart Glasses", "Tablet Lite", "Charging Dock"]
COMPONENTS = ["Display module", "Battery pack", "Main board", "Camera module", "Speaker assy", "Front cover",
              "Hinge assy", "Lens pair", "Charging cable", "Packaging kit", "Fan module", "Strap assy",
              "Antenna flex", "Button set", "Foam gasket"]


def _book(sheets):
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as xw:
        for name, (df, startrow) in sheets.items():
            df.to_excel(xw, sheet_name=name, index=False, startrow=startrow)
    return bio.getvalue()


def make_inputs(seed=7, n_parts=140):
    """Return (ctb_xlsx_bytes, mrp_xlsx_bytes, new_product_pns) shaped like the real inputs."""
    rng = np.random.default_rng(seed)
    run = dt.date.today()
    rows, cmap, mrp = [], [], []
    new_pns = set()
    for i in range(n_parts):
        prod = PRODUCTS[i % len(PRODUCTS)]
        pn = f"DM-{10000 + i * 7}-{rng.integers(1, 4):02d}"
        desc = f"{rng.choice(COMPONENTS)}, {prod}".upper()
        unit = float(np.round(rng.lognormal(3.0, 0.9), 2))
        qty = int(rng.integers(0, 900))
        ptype = rng.choice(["Company Owned", "Company Owned", "Company Owned", "Consigned", "Harvested"])
        suffix = {"Consigned": "-f", "Harvested": "-hb"}.get(ptype, "")
        bucket = rng.random()
        is_new = prod == "Smart Glasses" and bucket > 0.8 and ptype == "Company Owned"
        first_rec = run - dt.timedelta(days=int(rng.integers(5, 60) if is_new else rng.integers(120, 700)))
        last_act = run - dt.timedelta(days=int(rng.integers(0, 400)))
        rows.append({"PART_NO": pn + suffix, "PART_DESC": desc, "QTY_ONHAND": qty, "INV_VALUE": unit,
                     "EXT_COST": round(qty * unit, 2), "Family": "", "Part Type": ptype,
                     "LAST_ACTIVITY": last_act, "REC_DATE": first_rec})
        cmap.append({"PART_NO": pn + suffix, "Family": f"{i % len(PRODUCTS) + 1:02d} {prod.upper()}"})
        # demand profile: some parts healthy, some demand-drop, some zero, some missing from MRP
        if is_new:
            new_pns.add(pn)                                  # new-program seed stock
        if bucket < 0.07:
            continue                                         # no MRP row -> "unverified" guardrail
        if bucket < 0.30:
            dem = 0                                          # no demand
        elif bucket < 0.60:
            dem = int(qty * rng.uniform(0.1, 0.6))           # demand dropped below stock
        else:
            dem = int(qty * rng.uniform(0.9, 2.0))           # healthy
        mrp.append({"Component": pn, "Product Family": prod, "Company Owned": qty, "12 Week Demand": dem})
    # a few consumables (MRO pattern) so the carve-out shows up
    for j in range(6):
        pn = f"S-{500 + j}"
        q = int(rng.integers(50, 2000)); u = float(np.round(rng.uniform(0.2, 4), 2))
        rows.append({"PART_NO": pn, "PART_DESC": f"CONSUMABLE SUPPLY ITEM {j + 1}", "QTY_ONHAND": q,
                     "INV_VALUE": u, "EXT_COST": round(q * u, 2), "Family": "", "Part Type": "Company Owned",
                     "LAST_ACTIVITY": run, "REC_DATE": run - dt.timedelta(days=200)})
    ctb = _book({"Inventory Table": (pd.DataFrame(rows), 0), "Component Mapping": (pd.DataFrame(cmap), 0)})
    mrp_df = pd.DataFrame(mrp)[["Component", "Product Family", "Company Owned", "12 Week Demand"]]
    mrp_x = _book({"MRP Purchasing 12Weeks": (mrp_df, 1)})   # real file has its header on row 2
    return ctb, mrp_x, new_pns


def make_snapshots(seed=11, cycles=6):
    """Six demand-plan publications; each forecasts the same target month, so a falling line = demand cut."""
    rng = np.random.default_rng(seed)
    today = dt.date.today()
    target = (today.year + (today.month // 12), today.month % 12 + 1)   # next month
    snaps = []
    base = {p: int(rng.integers(800, 4000)) for p in PRODUCTS}
    trend = {p: rng.uniform(0.55, 1.1) for p in PRODUCTS}               # <1 = demand cut over cycles
    for c in range(cycles):
        f = c / (cycles - 1)
        vals = {p: max(0, int(base[p] * (1 + (trend[p] - 1) * f) * rng.uniform(0.95, 1.05))) for p in PRODUCTS}
        df = pd.DataFrame({"Model": list(vals), "M": list(vals.values())})
        mo = (today.month - (cycles - 1 - c) - 1) % 12 + 1
        snaps.append({"label": f"Plan {dt.date(2000, mo, 1):%b}", "mtime": c, "mdl": "Model",
                      "months": {"M": target}, "df": df})
    return snaps
