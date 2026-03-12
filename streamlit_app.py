import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO
from pathlib import Path
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

st.set_page_config(page_title="ShelfIQ 911", layout="wide")

# ---------------------------------------------------------
# BRANDING
# ---------------------------------------------------------

APP_TITLE = "ShelfIQ 911"
APP_SUBTITLE = "Retail Analytics & Shelf Optimization"
LOGO_PATH = "logo.png"


def render_header():
    col1, col2 = st.columns([1, 6])

    with col1:
        if Path(LOGO_PATH).exists():
            st.image(LOGO_PATH, width=120)

    with col2:
        st.markdown(f"# {APP_TITLE}")
        st.caption(APP_SUBTITLE)


# ---------------------------------------------------------
# REGION AUTO MAPPING
# ---------------------------------------------------------

STATE_REGION = {

"CT":"Northeast","ME":"Northeast","MA":"Northeast","NH":"Northeast","RI":"Northeast","VT":"Northeast",
"NJ":"Northeast","NY":"Northeast","PA":"Northeast",

"IL":"Midwest","IN":"Midwest","MI":"Midwest","OH":"Midwest","WI":"Midwest",
"IA":"Midwest","KS":"Midwest","MN":"Midwest","MO":"Midwest","NE":"Midwest",
"ND":"Midwest","SD":"Midwest",

"AL":"Southeast","AR":"Southeast","DE":"Southeast","FL":"Southeast","GA":"Southeast",
"KY":"Southeast","LA":"Southeast","MD":"Southeast","MS":"Southeast","NC":"Southeast",
"SC":"Southeast","TN":"Southeast","VA":"Southeast","WV":"Southeast",

"AZ":"Southwest","NM":"Southwest","OK":"Southwest","TX":"Southwest",

"CA":"West","CO":"West","ID":"West","MT":"West","NV":"West","OR":"West",
"UT":"West","WA":"West","WY":"West","AK":"West","HI":"West"

}


# ---------------------------------------------------------
# DATA CLEANING
# ---------------------------------------------------------

def normalize(df):
    df.columns = df.columns.str.lower().str.replace(" ","_")
    return df


# ---------------------------------------------------------
# DATA QUALITY CHECK
# ---------------------------------------------------------

def data_quality(sales):

    checks = []

    missing_store = sales["store_id"].isna().sum()
    checks.append(["Missing store_id",missing_store])

    missing_sku = sales["sku_id"].isna().sum()
    checks.append(["Missing sku_id",missing_sku])

    negative_units = (sales["units"] < 0).sum()
    checks.append(["Negative Units",negative_units])

    return pd.DataFrame(checks,columns=["check","count"])


# ---------------------------------------------------------
# KPI CALCULATIONS
# ---------------------------------------------------------

def calculate_metrics(sales):

    store_sales = sales.groupby("store_id")["sales_dollars"].sum().reset_index()

    expected = store_sales["sales_dollars"].mean()

    store_sales["spi"] = (store_sales["sales_dollars"] / expected) * 100

    under = store_sales[store_sales["spi"] < 80]

    revenue_gap = (expected - store_sales["sales_dollars"]).clip(lower=0)

    revenue_opportunity = revenue_gap.sum()

    return store_sales, under, revenue_opportunity


# ---------------------------------------------------------
# PDF REPORT
# ---------------------------------------------------------

def create_pdf(summary):

    buffer = BytesIO()
    c = canvas.Canvas(buffer,pagesize=letter)

    y = 750

    if Path(LOGO_PATH).exists():
        logo = ImageReader(LOGO_PATH)
        c.drawImage(logo,40,y,width=80)

    c.setFont("Helvetica-Bold",18)
    c.drawString(150,y,"ShelfIQ 911 Executive Report")

    y -= 60

    c.setFont("Helvetica",12)

    for k,v in summary.items():
        c.drawString(40,y,f"{k}: {v}")
        y -= 20

    c.save()

    buffer.seek(0)

    return buffer


# ---------------------------------------------------------
# APP
# ---------------------------------------------------------

render_header()

st.divider()

st.subheader("Upload Retail Data")

uploaded = st.file_uploader(
"Upload Excel file containing Sales_History, Products, Stores",
type=["xlsx"]
)

if uploaded:

    sales = normalize(pd.read_excel(uploaded,"Sales_History"))
    products = normalize(pd.read_excel(uploaded,"Products"))
    stores = normalize(pd.read_excel(uploaded,"Stores"))

    stores["region"] = stores["state"].map(STATE_REGION)

    quality = data_quality(sales)

    store_perf, underperf, revenue_opportunity = calculate_metrics(sales)

    avg_spi = store_perf["spi"].mean()

    health_score = round((avg_spi / 120) * 100 ,1)

    summary = {
        "Retail Health Score":health_score,
        "Avg SPI":round(avg_spi,1),
        "Underperforming Stores":len(underperf),
        "Revenue Opportunity":f"${int(revenue_opportunity):,}"
    }

    st.divider()

# ---------------------------------------------------------
# KPI CARDS
# ---------------------------------------------------------

    c1,c2,c3,c4 = st.columns(4)

    c1.metric("Retail Health",health_score)
    c2.metric("Avg SPI",round(avg_spi,1))
    c3.metric("Underperforming Stores",len(underperf))
    c4.metric("Revenue Opportunity",f"${int(revenue_opportunity):,}")

# ---------------------------------------------------------
# CHARTS
# ---------------------------------------------------------

    st.divider()

    st.subheader("Top Store Opportunities")

    top = store_perf.sort_values("spi").head(10)

    fig,ax = plt.subplots()

    ax.bar(top["store_id"].astype(str),top["spi"])

    ax.set_title("Lowest Store Performance")

    st.pyplot(fig)

# ---------------------------------------------------------
# TABS
# ---------------------------------------------------------

    tabs = st.tabs([
    "Executive Summary",
    "Data Quality",
    "Underperforming Stores"
    ])

# EXECUTIVE

    with tabs[0]:

        st.info(
        f"""
Retail Health Score is **{health_score}**.

There are **{len(underperf)} underperforming stores**.

Estimated revenue opportunity is **${int(revenue_opportunity):,}**.

Focus on stores with low SPI and brands with weak velocity.
"""
        )

# DATA QUALITY

    with tabs[1]:

        st.dataframe(quality,use_container_width=True)

# UNDERPERFORMING

    with tabs[2]:

        st.dataframe(
        underperf.sort_values("spi"),
        use_container_width=True
        )

# ---------------------------------------------------------
# DOWNLOADS
# ---------------------------------------------------------

    st.divider()

    pdf = create_pdf(summary)

    st.download_button(
    "Download Executive PDF",
    pdf,
    "shelfiq_report.pdf"
    )

    output = BytesIO()

    with pd.ExcelWriter(output,engine="openpyxl") as writer:
        store_perf.to_excel(writer,"Store_Performance",index=False)
        underperf.to_excel(writer,"Underperforming",index=False)

    st.download_button(
    "Download Excel Results",
    output.getvalue(),
    "shelfiq_results.xlsx"
    )