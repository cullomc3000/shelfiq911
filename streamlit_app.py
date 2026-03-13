import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

st.set_page_config(
    page_title="ShelfIQ 911",
    layout="wide"
)

# ---------------------------
# STYLE
# ---------------------------

st.markdown("""
<style>

.metric-card{
padding:18px;
border-radius:14px;
color:white;
font-weight:600;
box-shadow:0px 6px 18px rgba(0,0,0,0.2);
}

.blue{background:linear-gradient(135deg,#2563eb,#3b82f6);}
.purple{background:linear-gradient(135deg,#7c3aed,#9333ea);}
.teal{background:linear-gradient(135deg,#0f766e,#14b8a6);}
.orange{background:linear-gradient(135deg,#c2410c,#f97316);}
.dark{background:linear-gradient(135deg,#111827,#374151);}

.panel{
background:white;
padding:18px;
border-radius:14px;
box-shadow:0px 6px 18px rgba(0,0,0,0.1);
}

.story-box{
background:linear-gradient(135deg,#1e3a8a,#2563eb);
color:white;
padding:22px;
border-radius:18px;
font-size:18px;
}

.action-box{
background:white;
border-left:6px solid #2563eb;
padding:14px;
border-radius:10px;
margin-bottom:10px;
}

</style>
""", unsafe_allow_html=True)

# ---------------------------
# HEADER
# ---------------------------

st.title("ShelfIQ 911")
st.caption("Retail Analytics • Distribution Intelligence • Shelf Optimization")

# ---------------------------
# SAMPLE DATA
# ---------------------------

np.random.seed(1)

stores = 120
skus = 40

df = pd.DataFrame({
    "store":np.random.randint(1,stores,500),
    "sku":np.random.randint(1,skus,500),
    "sales":np.random.randint(50,500,500),
    "week":np.random.randint(1,52,500)
})

# ---------------------------
# KPI CARDS
# ---------------------------

k1,k2,k3,k4,k5 = st.columns(5)

with k1:
    st.markdown('<div class="metric-card blue">Retail Health<br><h2>87</h2></div>',unsafe_allow_html=True)

with k2:
    st.markdown('<div class="metric-card purple">Data Quality<br><h2>94%</h2></div>',unsafe_allow_html=True)

with k3:
    st.markdown('<div class="metric-card teal">Revenue Opportunity<br><h2>$1.3M</h2></div>',unsafe_allow_html=True)

with k4:
    st.markdown('<div class="metric-card orange">Distribution Gap<br><h2>18%</h2></div>',unsafe_allow_html=True)

with k5:
    st.markdown('<div class="metric-card dark">Return Impact<br><h2>3.1%</h2></div>',unsafe_allow_html=True)

st.markdown("---")

# ---------------------------
# MCKINSEY STORY GENERATOR
# ---------------------------

st.markdown("## Executive Dashboard")

sales_total = df.sales.sum()
avg_sales = df.sales.mean()

story = f"""
Retail performance is **stable but under-optimized**.

Analysis indicates approximately **$1.3M in unrealized revenue** driven by:

• distribution gaps in key retail partners 
• underperforming store execution 
• shelf productivity imbalance 

Immediate actions should focus on expanding high-velocity SKUs and correcting store execution issues.
"""

st.markdown(f'<div class="story-box">{story}</div>',unsafe_allow_html=True)

st.markdown("")

# ---------------------------
# CHARTS
# ---------------------------

c1,c2,c3 = st.columns(3)

with c1:
    fig = px.bar(
        df.groupby("sku").sales.sum().reset_index(),
        x="sku",
        y="sales",
        title="SKU Velocity"
    )
    st.plotly_chart(fig,use_container_width=True)

with c2:
    fig = px.bar(
        df.groupby("store").sales.sum().reset_index(),
        x="store",
        y="sales",
        title="Store Performance"
    )
    st.plotly_chart(fig,use_container_width=True)

with c3:
    fig = px.histogram(
        df,
        x="sales",
        nbins=20,
        title="Sales Distribution"
    )
    st.plotly_chart(fig,use_container_width=True)

# ---------------------------
# ACTION CENTER
# ---------------------------

st.markdown("## Strategic Action Center")

a1,a2 = st.columns(2)

with a1:

    st.markdown('<div class="action-box"><b>Distribution Expansion</b><br>Expand SKU 12 into 48 additional stores</div>',unsafe_allow_html=True)

    st.markdown('<div class="action-box"><b>Store Execution</b><br>Investigate store cluster underperforming by 32%</div>',unsafe_allow_html=True)

    st.markdown('<div class="action-box"><b>Momentum Opportunity</b><br>SKU 8 trending upward +22%</div>',unsafe_allow_html=True)

with a2:

    st.markdown('<div class="action-box"><b>Shelf Optimization</b><br>Increase facings for top 5 SKUs</div>',unsafe_allow_html=True)

    st.markdown('<div class="action-box"><b>Assortment Fix</b><br>Remove low velocity SKU 3</div>',unsafe_allow_html=True)

    st.markdown('<div class="action-box"><b>Retailer Sell-In</b><br>Pitch distribution expansion to Target</div>',unsafe_allow_html=True)

st.markdown("---")

# ---------------------------
# DATA TABLE
# ---------------------------

st.markdown("## Detailed Data")

st.dataframe(df)
