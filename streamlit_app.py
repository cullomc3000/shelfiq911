
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import plotly.express as px

st.set_page_config(page_title="ShelfIQ 911", layout="wide")

st.title("ShelfIQ 911")
st.caption("Retail Analytics • Distribution Intelligence • Shelf Optimization")

uploaded_file = st.file_uploader(
    "Upload Client Excel File (Sheets: Sales_History, Products, Stores, Shelf_Snapshot optional)",
    type=["xlsx"]
)

if uploaded_file:

    sales = pd.read_excel(uploaded_file, sheet_name="Sales_History")
    products = pd.read_excel(uploaded_file, sheet_name="Products")
    stores = pd.read_excel(uploaded_file, sheet_name="Stores")

    try:
        shelf = pd.read_excel(uploaded_file, sheet_name="Shelf_Snapshot")
    except:
        shelf = pd.DataFrame()

    st.success("File Loaded Successfully")

    sales.columns = sales.columns.str.lower()
    products.columns = products.columns.str.lower()
    stores.columns = stores.columns.str.lower()

    sales['week_end_date'] = pd.to_datetime(sales['week_end_date'], errors='coerce')
    sales['units'] = pd.to_numeric(sales['units'], errors='coerce').fillna(0)
    sales['sales_dollars'] = pd.to_numeric(sales['sales_dollars'], errors='coerce').fillna(0)

    sku_velocity = (
        sales.groupby("sku_id")
        .agg(total_units=("units","sum"),
             total_sales=("sales_dollars","sum"),
             stores=("store_id","nunique"))
        .reset_index()
    )

    sku_velocity["velocity"] = sku_velocity["total_units"] / sku_velocity["stores"]

    store_perf = (
        sales.groupby("store_id")
        .agg(sales=("sales_dollars","sum"))
        .reset_index()
    )

    avg_sales = store_perf["sales"].mean()
    store_perf["spi"] = (store_perf["sales"] / avg_sales) * 100
    store_perf["opportunity"] = np.where(store_perf["spi"] < 80, avg_sales - store_perf["sales"], 0)

    distribution = (
        sales.groupby(["sku_id","store_id"])
        .size()
        .reset_index(name="present")
    )

    dist_gap = (
        distribution.groupby("sku_id")
        .agg(stores=("store_id","nunique"))
        .reset_index()
    )

    total_stores = stores["store_id"].nunique()
    dist_gap["distribution_gap"] = total_stores - dist_gap["stores"]

    st.header("Executive Dashboard")

    c1,c2,c3,c4 = st.columns(4)

    with c1:
        st.metric("Stores", stores["store_id"].nunique())

    with c2:
        st.metric("SKUs", products["sku_id"].nunique())

    with c3:
        st.metric("Revenue Opportunity", round(store_perf["opportunity"].sum(),2))

    with c4:
        st.metric("Avg Store Performance", round(store_perf["spi"].mean(),1))

    st.subheader("Top Store Opportunities")

    fig = px.bar(
        store_perf.sort_values("opportunity",ascending=False).head(10),
        x="store_id",
        y="opportunity"
    )

    st.plotly_chart(fig,use_container_width=True)

    st.subheader("Top SKU Velocity")

    fig2 = px.bar(
        sku_velocity.sort_values("velocity",ascending=False).head(10),
        x="sku_id",
        y="velocity"
    )

    st.plotly_chart(fig2,use_container_width=True)

    st.subheader("Distribution Gaps")

    fig3 = px.bar(
        dist_gap.sort_values("distribution_gap",ascending=False).head(10),
        x="sku_id",
        y="distribution_gap"
    )

    st.plotly_chart(fig3,use_container_width=True)

    st.subheader("Data Tables")

    tab1,tab2,tab3 = st.tabs(["Store Performance","SKU Velocity","Distribution"])

    with tab1:
        st.dataframe(store_perf)

    with tab2:
        st.dataframe(sku_velocity)

    with tab3:
        st.dataframe(dist_gap)
