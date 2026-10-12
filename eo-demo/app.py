"""
E&O (Excess & Obsolete) Automation - public demo.
Runs on generated sample data by default; you can also upload your own files in the same layout.
Run locally:  python -m streamlit run app.py
"""
import datetime as dt
import streamlit as st
import eo_engine as E
import sample_data as S

st.set_page_config(page_title="E&O Automation - Demo", layout="wide")

st.title("E&O Automation - demo")
st.caption("Demo version with synthetic sample data - no real parts, customers, or prices. "
           "Built by Ankit Kumar (Python, pandas, Streamlit, openpyxl).")

with st.expander("What this tool does", expanded=True):
    st.markdown(
        "- **Problem:** the quarterly excess & obsolete (E&O) inventory review was assembled by hand from a "
        "clear-to-build (CTB) inventory report and an MRP demand file.\n"
        "- **What it automates:** reads both files, keeps company-owned stock only, carves out consumables and "
        "new-product seed stock, computes excess after 13 weeks of demand, drafts a disposition and reason per part, "
        "flags parts whose demand link can't be verified, and builds the customer-ready Excel workbook.\n"
        "- **Demand waterfall:** shows how each product's forecast for the same month fell across plan cycles - "
        "the demand cut that created the excess.\n"
        "- **Human in the loop:** dispositions and reasons are drafts for planner approval.")

src = st.radio("Input data", ["Use generated sample data", "Upload my own files"], horizontal=True)
ctb_b = mrp_b = None
new_pns = set()
if src.startswith("Use"):
    seed = st.number_input("Sample-data seed (change to regenerate)", value=7, step=1)
    ctb_b, mrp_b, new_pns = S.make_inputs(seed=int(seed))
    c1, c2 = st.columns(2)
    c1.download_button("Download sample CTB input", ctb_b, "sample_CTB.xlsx")
    c2.download_button("Download sample MRP input", mrp_b, "sample_MRP.xlsx")
else:
    c1, c2 = st.columns(2)
    f1 = c1.file_uploader("CTB workbook (sheet 'Inventory Table')", type=["xlsx"])
    f2 = c2.file_uploader("MRP workbook (sheet 'MRP Purchasing 12Weeks')", type=["xlsx"])
    ctb_b = f1.getvalue() if f1 else None
    mrp_b = f2.getvalue() if f2 else None

p1, p2, p3, p4 = st.columns(4)
site = p1.selectbox("Site", ["SITE-1", "SITE-2", "SITE-3"])
run_date = p2.date_input("Run date", dt.date.today())
markup = p3.number_input("Markup %", value=5.0, step=0.5) / 100.0
materiality = p4.number_input("Materiality $ (min billable)", value=50.0, step=10.0)

if st.button("Run E&O analysis", type="primary", disabled=not (ctb_b and mrp_b)):
    inv, meta = E.load_inventory(ctb_b)
    dem, _ = E.load_demand(mrp_b)
    eligible, cons_df, seed_df = E.split_eligible(inv, set(), new_pns, run_date, seed_days=90)
    eo = E.run_eo(eligible, dem, markup=markup, materiality=materiality)
    k = E.kpis(eo)

    st.success(f"{len(eo)} E&O lines  |  {meta['parts']:,} owned parts read; "
               f"{len(cons_df)} consumable and {len(seed_df)} seed-stock parts carved out")
    m = st.columns(5)
    m[0].metric("E&O parts", f"{k['parts']:,}")
    m[1].metric("Total excess $", f"${k['eoh_dollars']:,.0f}")
    m[2].metric(f"Billable $ (w/{markup:.0%})", f"${k['billable']:,.0f}")
    m[3].metric("Needs review (unverified)", f"{k['unverified']:,}")
    m[4].metric("Carved out $", f"${cons_df['IOH_DOLLARS'].sum() + seed_df['IOH_DOLLARS'].sum():,.0f}")

    if len(eo):
        a, b = st.columns(2)
        a.subheader("Billable excess by product")
        a.bar_chart(eo.groupby("PRODUCT")["EOH_BILL"].sum().sort_values(ascending=False))
        b.subheader("Billable excess by cause")
        b.bar_chart(eo.groupby("CAUSE")["EOH_BILL"].sum())
        st.subheader("E&O working table")
        st.dataframe(eo[["PRODUCT", "PART_NO", "PART_DESC", "IOH_QTY", "DEM13", "EOH_QTY", "EOH_BILL",
                         "REQUEST", "CONFIDENCE", "REASON"]], width="stretch")

    with st.expander(f"Carved out: consumables ({len(cons_df)}) and seed stock ({len(seed_df)})"):
        st.dataframe(cons_df[["PART_NO", "PART_DESC", "IOH_QTY", "IOH_DOLLARS", "REASON"]], width="stretch")
        st.dataframe(seed_df[["PART_NO", "PART_DESC", "IOH_QTY", "IOH_DOLLARS", "REASON"]], width="stretch")

    wf = E.build_waterfall(S.make_snapshots()) if src.startswith("Use") else None
    if wf:
        st.subheader(f"Demand waterfall - forecast for {wf['target']} across plan cycles")
        st.caption("A falling line = the customer cut demand after stock was bought against the earlier plan.")
        import altair as alt
        tbl = wf["table"].copy(); tbl.index.name = "Product"
        long = tbl.reset_index().melt(id_vars="Product", var_name="Cycle", value_name="Units")
        st.altair_chart(alt.Chart(long).mark_line(point=True).encode(
            x=alt.X("Cycle:N", sort=list(tbl.columns), title="Plan cycle (older -> current)"),
            y=alt.Y("Units:Q", title="Forecast units"),
            color="Product:N",
            tooltip=["Product:N", "Cycle:N", alt.Tooltip("Units:Q", format=",")]).properties(height=380),
            width="stretch")

    st.download_button("Download E&O workbook (Excel)",
                       data=E.build_workbook(eo, site, run_date, wf, consumables=cons_df, seed=seed_df),
                       file_name=f"E&O Update - {site} - {run_date}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Pick sample data (or upload files) and click **Run E&O analysis**.")
