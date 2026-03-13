
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(
    page_title="ShelfIQ 911",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.block-container{
    padding-top:1rem;
    padding-bottom:2rem;
    max-width:1500px;
}
.stApp{
    background:linear-gradient(180deg,#f6f9ff 0%,#eef4ff 100%);
}
.metric-card{
    padding:18px;
    border-radius:18px;
    color:white;
    font-weight:600;
    box-shadow:0px 8px 22px rgba(0,0,0,0.18);
    min-height:120px;
}
.metric-label{
    font-size:0.9rem;
    opacity:0.9;
    margin-bottom:6px;
}
.metric-value{
    font-size:2rem;
    font-weight:800;
    line-height:1.05;
}
.metric-sub{
    font-size:0.82rem;
    opacity:0.88;
    margin-top:6px;
}
.blue{background:linear-gradient(135deg,#2563eb,#3b82f6);}
.purple{background:linear-gradient(135deg,#7c3aed,#9333ea);}
.teal{background:linear-gradient(135deg,#0f766e,#14b8a6);}
.orange{background:linear-gradient(135deg,#c2410c,#f97316);}
.dark{background:linear-gradient(135deg,#111827,#374151);}
.panel{
    background:white;
    padding:18px;
    border-radius:18px;
    box-shadow:0px 6px 18px rgba(0,0,0,0.08);
    border:1px solid #dbeafe;
}
.story-box{
    background:linear-gradient(135deg,#1e3a8a,#2563eb);
    color:white;
    padding:22px;
    border-radius:18px;
    box-shadow:0px 10px 24px rgba(37,99,235,0.22);
}
.story-head{
    font-size:1.05rem;
    font-weight:800;
    margin-bottom:10px;
}
.story-body{
    font-size:1rem;
    line-height:1.6;
}
.story-card{
    background:white;
    border:1px solid #dbeafe;
    border-radius:14px;
    padding:14px;
    min-height:125px;
    box-shadow:0px 6px 16px rgba(37,99,235,0.06);
}
.story-card-title{
    color:#2563eb;
    font-size:0.9rem;
    font-weight:800;
    margin-bottom:6px;
}
.story-card-body{
    color:#0f172a;
    font-size:0.9rem;
    line-height:1.5;
}
.action-box{
    background:white;
    border-left:6px solid #2563eb;
    padding:14px;
    border-radius:12px;
    margin-bottom:10px;
    box-shadow:0px 6px 16px rgba(0,0,0,0.06);
}
.action-box.high{border-left-color:#dc2626;}
.action-box.medium{border-left-color:#f59e0b;}
.action-box.low{border-left-color:#10b981;}
.sidebar-title{
    color:white;
    font-size:1.1rem;
    font-weight:800;
}
[data-testid="stSidebar"]{
    background:linear-gradient(180deg,#312e81 0%,#1d4ed8 100%);
}
[data-testid="stSidebar"] *{
    color:white;
}
</style>
""", unsafe_allow_html=True)

STATE_TO_REGION = {
    "CT": "Northeast", "ME": "Northeast", "MA": "Northeast", "NH": "Northeast",
    "RI": "Northeast", "VT": "Northeast", "NJ": "Northeast", "NY": "Northeast", "PA": "Northeast",
    "IL": "Midwest", "IN": "Midwest", "MI": "Midwest", "OH": "Midwest", "WI": "Midwest",
    "IA": "Midwest", "KS": "Midwest", "MN": "Midwest", "MO": "Midwest", "NE": "Midwest",
    "ND": "Midwest", "SD": "Midwest",
    "AL": "Southeast", "AR": "Southeast", "DE": "Southeast", "DC": "Southeast", "FL": "Southeast",
    "GA": "Southeast", "KY": "Southeast", "LA": "Southeast", "MD": "Southeast", "MS": "Southeast",
    "NC": "Southeast", "SC": "Southeast", "TN": "Southeast", "VA": "Southeast", "WV": "Southeast",
    "AZ": "Southwest", "NM": "Southwest", "OK": "Southwest", "TX": "Southwest",
    "AK": "West", "CA": "West", "CO": "West", "HI": "West", "ID": "West", "MT": "West",
    "NV": "West", "OR": "West", "UT": "West", "WA": "West", "WY": "West"
}

def metric_card_html(label, value, sub, klass):
    return f'''
    <div class="metric-card {klass}">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        <div class="metric-sub">{sub}</div>
    </div>
    '''

def normalize_columns(df):
    out = df.copy()
    out.columns = [str(c).strip().lower().replace(" ", "_").replace("-", "_") for c in out.columns]
    return out

def format_currency(x):
    try:
        return f"${float(x):,.0f}"
    except Exception:
        return "$0"

def build_story(summary):
    headline = (
        f"Retail performance is {summary['health_label'].lower()} with "
        f"{format_currency(summary['revenue_opportunity'])} in modeled upside and "
        f"{int(summary['underperforming_stores'])} underperforming stores."
    )
    so_what = (
        "The largest value pools are store execution, distribution expansion, "
        "and targeted shelf-space optimization."
    )
    risk = (
        f"Data quality score is {summary['data_quality_score']:.1f}. "
        f"Returns impact is {summary['return_impact_pct']:.1f}% of units."
    )
    evidence = [
        f"Average Store Performance Index is {summary['avg_spi']:.1f}.",
        f"Distribution gap averages {summary['avg_distribution_gap_pct']:.1f}% across tracked SKUs.",
        "Top momentum SKUs are showing stronger recent velocity than their long-term baseline.",
        "Highest upside comes from fixing underperforming stores before broad expansion."
    ]
    return {"headline": headline, "so_what": so_what, "risk": risk, "evidence": evidence}

def classify_health(score):
    if score >= 90:
        return "Excellent"
    if score >= 75:
        return "Strong"
    if score >= 60:
        return "Fair"
    return "Weak"

def safe_read_excel(uploaded_file, sheet_name):
    uploaded_file.seek(0)
    return pd.read_excel(uploaded_file, sheet_name=sheet_name)

def generate_demo_data():
    np.random.seed(7)
    retailers = ["Walmart", "Kroger", "Home Depot", "Walgreens", "7-Eleven"]
    states = ["GA", "TX", "FL", "NC", "CA", "OH", "PA", "AZ", "TN", "VA"]
    categories = ["Beverage", "Snacks", "Hardware", "Health", "Convenience"]
    brands = ["BrandA", "BrandB", "BrandC", "BrandD", "BrandE", "BrandF"]

    stores = []
    store_id = 1001
    for retailer in retailers:
        for _ in range(12):
            state = np.random.choice(states)
            stores.append({
                "store_id": str(store_id),
                "retailer": retailer,
                "state": state,
                "format": np.random.choice(["Urban", "Suburban", "Neighborhood", "Large Format"])
            })
            store_id += 1
    stores_df = pd.DataFrame(stores)

    products = []
    for i in range(1, 31):
        products.append({
            "sku_id": f"SKU{i:03d}",
            "brand": np.random.choice(brands),
            "category": np.random.choice(categories)
        })
    products_df = pd.DataFrame(products)

    dates = pd.date_range("2024-01-07", "2025-12-28", freq="W-SUN")
    sales_rows = []
    for _, s in stores_df.iterrows():
        active_skus = np.random.choice(products_df["sku_id"], size=np.random.randint(10, 18), replace=False)
        store_factor = np.random.uniform(0.7, 1.35)
        for sku in active_skus:
            base_units = np.random.randint(6, 24)
            trend = np.random.uniform(0.92, 1.12)
            price = np.random.uniform(4.5, 28)
            for n, dt in enumerate(dates):
                seasonality = 1 + 0.15 * np.sin(n / 6)
                momentum_bump = 1.0
                if sku in ["SKU001", "SKU005", "SKU010"]:
                    momentum_bump = 1.20 if dt.year == 2025 else 0.92
                if sku in ["SKU003", "SKU008"]:
                    momentum_bump = 0.78 if dt.year == 2025 else 1.05
                units = max(0, int(round(base_units * store_factor * seasonality * trend * momentum_bump + np.random.normal(0, 2))))
                if np.random.rand() < 0.01:
                    units = -np.random.randint(1, 4)
                sales_rows.append({
                    "store_id": s["store_id"],
                    "sku_id": sku,
                    "week_end_date": dt,
                    "units": units,
                    "sales_dollars": round(units * price, 2)
                })
    sales_df = pd.DataFrame(sales_rows)

    shelf_rows = []
    sampled = sales_df.groupby(["store_id", "sku_id"], as_index=False)["units"].sum().head(250)
    for _, r in sampled.iterrows():
        shelf_rows.append({
            "store_id": r["store_id"],
            "sku_id": r["sku_id"],
            "facings": np.random.randint(1, 8),
            "shelf_share": round(np.random.uniform(0.02, 0.22), 3)
        })
    shelf_df = pd.DataFrame(shelf_rows)
    return sales_df, products_df, stores_df, shelf_df

def run_engine(sales, products, stores, shelf):
    sales = normalize_columns(sales)
    products = normalize_columns(products)
    stores = normalize_columns(stores)
    shelf = normalize_columns(shelf) if shelf is not None and len(shelf) > 0 else pd.DataFrame()

    required_sales = {"store_id", "sku_id", "week_end_date", "units"}
    required_products = {"sku_id"}
    required_stores = {"store_id", "retailer", "state"}
    missing = {}
    if not required_sales.issubset(sales.columns):
        missing["Sales_History"] = sorted(list(required_sales - set(sales.columns)))
    if not required_products.issubset(products.columns):
        missing["Products"] = sorted(list(required_products - set(products.columns)))
    if not required_stores.issubset(stores.columns):
        missing["Stores"] = sorted(list(required_stores - set(stores.columns)))
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    if "sales_dollars" not in sales.columns:
        sales["sales_dollars"] = 0.0
    if "brand" not in products.columns:
        products["brand"] = "Unknown"
    if "category" not in products.columns:
        products["category"] = "Unknown"
    if "format" not in stores.columns:
        stores["format"] = "Unknown"

    sales["store_id"] = sales["store_id"].astype(str).str.strip()
    sales["sku_id"] = sales["sku_id"].astype(str).str.strip()
    products["sku_id"] = products["sku_id"].astype(str).str.strip()
    stores["store_id"] = stores["store_id"].astype(str).str.strip()
    stores["state"] = stores["state"].astype(str).str.upper().str.strip()
    sales["week_end_date"] = pd.to_datetime(sales["week_end_date"], errors="coerce")
    sales["units"] = pd.to_numeric(sales["units"], errors="coerce").fillna(0)
    sales["sales_dollars"] = pd.to_numeric(sales["sales_dollars"], errors="coerce").fillna(0)
    stores["region"] = stores["state"].map(STATE_TO_REGION).fillna("Unknown")

    invalid_dates = sales["week_end_date"].isna().sum()
    negative_units_pct = ((sales["units"] < 0).sum() / max(len(sales), 1)) * 100
    unmatched_skus = (~sales["sku_id"].isin(products["sku_id"])).sum()
    unmatched_stores = (~sales["store_id"].isin(stores["store_id"])).sum()
    dq_penalty = invalid_dates * 2 + unmatched_skus * 0.5 + unmatched_stores * 0.5 + max(0, negative_units_pct - 5) * 2
    data_quality_score = max(0, round(100 - dq_penalty, 1))
    return_impact_pct = round((sales.loc[sales["units"] < 0, "units"].abs().sum() / max(sales["units"].abs().sum(), 1)) * 100, 2)
    sales = sales.dropna(subset=["week_end_date"]).copy()

    sales_enriched = sales.merge(products, on="sku_id", how="left").merge(stores, on="store_id", how="left")
    max_week = sales_enriched["week_end_date"].max()
    trailing_13_start = max_week - pd.Timedelta(weeks=13)
    trailing_52_start = max_week - pd.Timedelta(weeks=52)
    sales_13 = sales_enriched[sales_enriched["week_end_date"] > trailing_13_start].copy()
    sales_52 = sales_enriched[sales_enriched["week_end_date"] > trailing_52_start].copy()
    weeks_13 = max(sales_13["week_end_date"].nunique(), 1)
    weeks_52 = max(sales_52["week_end_date"].nunique(), 1)

    sku_velocity = (
        sales_13.groupby(["sku_id", "brand", "category"], dropna=False)
        .agg(total_units=("units", "sum"), total_sales=("sales_dollars", "sum"), active_stores=("store_id", "nunique"))
        .reset_index()
    )
    sku_velocity["velocity"] = sku_velocity["total_units"] / sku_velocity["active_stores"].clip(lower=1) / weeks_13
    cat_avg = sku_velocity.groupby("category", dropna=False)["velocity"].mean().rename("category_avg_velocity").reset_index()
    sku_velocity = sku_velocity.merge(cat_avg, on="category", how="left")
    sku_velocity["sku_velocity_index"] = (sku_velocity["velocity"] / sku_velocity["category_avg_velocity"].replace(0, np.nan)) * 100
    sku_velocity["sku_velocity_index"] = sku_velocity["sku_velocity_index"].fillna(0)

    store_perf = (
        sales_13.groupby(["store_id", "retailer", "region", "state", "format"], dropna=False)
        .agg(actual_sales=("sales_dollars", "sum")).reset_index()
    )
    peer_avg = (
        store_perf.groupby(["retailer", "format", "region"], dropna=False)["actual_sales"]
        .mean().rename("expected_sales").reset_index()
    )
    store_perf = store_perf.merge(peer_avg, on=["retailer", "format", "region"], how="left")
    store_perf["expected_sales"] = store_perf["expected_sales"].fillna(store_perf["actual_sales"].mean())
    store_perf["spi"] = (store_perf["actual_sales"] / store_perf["expected_sales"].replace(0, np.nan)) * 100
    store_perf["spi"] = store_perf["spi"].fillna(0)
    store_perf["sales_gap"] = store_perf["expected_sales"] - store_perf["actual_sales"]
    store_perf["revenue_opportunity"] = np.where(store_perf["sales_gap"] > 0, store_perf["sales_gap"], 0)
    store_perf["underperforming"] = store_perf["spi"] < 80
    underperf = store_perf[store_perf["underperforming"]].sort_values("revenue_opportunity", ascending=False).reset_index(drop=True)

    carried = (
        sales_13.groupby(["brand", "category", "retailer", "store_id"], dropna=False)
        .agg(total_units=("units", "sum")).reset_index()
    )
    carried = carried[carried["total_units"] > 0]
    retailer_store_universe = stores.groupby("retailer", dropna=False)["store_id"].nunique().rename("retailer_store_universe").reset_index()
    distribution_gap = (
        carried.groupby(["brand", "category", "retailer"], dropna=False)["store_id"]
        .nunique().rename("current_store_count").reset_index()
        .merge(retailer_store_universe, on="retailer", how="left")
    )
    distribution_gap["distribution_gap_count"] = (distribution_gap["retailer_store_universe"] - distribution_gap["current_store_count"]).clip(lower=0)
    distribution_gap["distribution_gap_pct"] = (distribution_gap["distribution_gap_count"] / distribution_gap["retailer_store_universe"].replace(0, np.nan)) * 100
    distribution_gap["distribution_gap_pct"] = distribution_gap["distribution_gap_pct"].fillna(0)

    sales_enriched["year"] = sales_enriched["week_end_date"].dt.year
    yearly = (
        sales_enriched.groupby(["sku_id", "brand", "category", "year"], dropna=False)
        .agg(yearly_sales=("sales_dollars", "sum"), yearly_units=("units", "sum")).reset_index()
    )
    years = sorted(yearly["year"].dropna().unique())
    yoy = pd.DataFrame()
    if len(years) >= 2:
        py = years[-2]
        cy = years[-1]
        prev_df = yearly[yearly["year"] == py][["sku_id", "yearly_sales", "yearly_units"]].rename(
            columns={"yearly_sales": f"sales_{py}", "yearly_units": f"units_{py}"}
        )
        curr_df = yearly[yearly["year"] == cy][["sku_id", "brand", "category", "yearly_sales", "yearly_units"]].rename(
            columns={"yearly_sales": f"sales_{cy}", "yearly_units": f"units_{cy}"}
        )
        yoy = curr_df.merge(prev_df, on="sku_id", how="left")
        yoy[f"sales_{py}"] = yoy[f"sales_{py}"].fillna(0)
        yoy[f"units_{py}"] = yoy[f"units_{py}"].fillna(0)
        yoy["yoy_sales_growth_pct"] = np.where(
            yoy[f"sales_{py}"] > 0,
            ((yoy[f"sales_{cy}"] - yoy[f"sales_{py}"]) / yoy[f"sales_{py}"]) * 100,
            np.nan
        )
        yoy["yoy_units_growth_pct"] = np.where(
            yoy[f"units_{py}"] > 0,
            ((yoy[f"units_{cy}"] - yoy[f"units_{py}"]) / yoy[f"units_{py}"]) * 100,
            np.nan
        )

    v13 = sales_13.groupby("sku_id", dropna=False).agg(units_13=("units", "sum"), stores_13=("store_id", "nunique")).reset_index()
    v13["velocity_13w"] = v13["units_13"] / v13["stores_13"].clip(lower=1) / weeks_13
    v52 = sales_52.groupby("sku_id", dropna=False).agg(units_52=("units", "sum"), stores_52=("store_id", "nunique")).reset_index()
    v52["velocity_52w"] = v52["units_52"] / v52["stores_52"].clip(lower=1) / weeks_52
    momentum = v13.merge(v52, on="sku_id", how="outer").merge(products[["sku_id", "brand", "category"]], on="sku_id", how="left")
    momentum["velocity_13w"] = momentum["velocity_13w"].fillna(0)
    momentum["velocity_52w"] = momentum["velocity_52w"].fillna(0)
    momentum["momentum_ratio"] = np.where(momentum["velocity_52w"] > 0, momentum["velocity_13w"] / momentum["velocity_52w"], np.nan)
    momentum["momentum_flag"] = np.select(
        [momentum["momentum_ratio"] >= 1.20, momentum["momentum_ratio"] <= 0.80],
        ["Trending Up", "Trending Down"],
        default="Stable"
    )

    shelf_metrics = pd.DataFrame()
    if shelf is not None and len(shelf) > 0 and {"store_id", "sku_id"}.issubset(shelf.columns):
        shelf["store_id"] = shelf["store_id"].astype(str).str.strip()
        shelf["sku_id"] = shelf["sku_id"].astype(str).str.strip()
        if "facings" not in shelf.columns:
            shelf["facings"] = np.nan
        shelf["facings"] = pd.to_numeric(shelf["facings"], errors="coerce")
        shelf_metrics = (
            shelf.merge(products, on="sku_id", how="left")
            .merge(
                sales_13.groupby(["store_id", "sku_id"], dropna=False)
                .agg(total_sales=("sales_dollars", "sum"), total_units=("units", "sum"))
                .reset_index(),
                on=["store_id", "sku_id"], how="left"
            )
            .merge(stores[["store_id", "retailer", "region"]], on="store_id", how="left")
        )
        shelf_metrics["total_sales"] = shelf_metrics["total_sales"].fillna(0)
        shelf_metrics["facings"] = pd.to_numeric(shelf_metrics["facings"], errors="coerce")
        shelf_metrics["sales_per_facing"] = (shelf_metrics["total_sales"] / shelf_metrics["facings"].replace(0, np.nan)).fillna(0)
        cat_spf = shelf_metrics.groupby("category", dropna=False)["sales_per_facing"].mean().rename("category_avg_sales_per_facing").reset_index()
        shelf_metrics = shelf_metrics.merge(cat_spf, on="category", how="left")
        shelf_metrics["space_efficiency_index"] = (
            shelf_metrics["sales_per_facing"] / shelf_metrics["category_avg_sales_per_facing"].replace(0, np.nan)
        ) * 100
        shelf_metrics["space_efficiency_index"] = shelf_metrics["space_efficiency_index"].fillna(0)
        shelf_metrics["shelf_action"] = np.select(
            [shelf_metrics["space_efficiency_index"] >= 120, shelf_metrics["space_efficiency_index"] < 80],
            ["Increase Facings", "Reduce / Review"],
            default="Hold"
        )

    avg_spi = float(store_perf["spi"].fillna(100).mean()) if len(store_perf) else 100
    underperf_rate = float(store_perf["underperforming"].mean()) if len(store_perf) else 0
    avg_dist_gap_pct = float(distribution_gap["distribution_gap_pct"].fillna(0).mean()) if len(distribution_gap) else 0
    health_score = round(
        (min(max(avg_spi, 0), 120) / 120) * 40 +
        (1 - min(max(underperf_rate, 0), 1)) * 20 +
        (1 - min(max(avg_dist_gap_pct / 100, 0), 1)) * 20 +
        (data_quality_score / 100) * 20, 1
    )

    summary = {
        "retail_health_score": health_score,
        "health_label": classify_health(health_score),
        "data_quality_score": data_quality_score,
        "store_count": int(stores["store_id"].nunique()),
        "sku_count": int(products["sku_id"].nunique()),
        "underperforming_stores": int(store_perf["underperforming"].sum()),
        "avg_spi": round(avg_spi, 2),
        "avg_distribution_gap_pct": round(avg_dist_gap_pct, 2),
        "revenue_opportunity": float(store_perf["revenue_opportunity"].sum()),
        "return_impact_pct": return_impact_pct
    }

    recommendations = []
    if len(underperf):
        r = underperf.iloc[0]
        recommendations.append({
            "priority": "High",
            "type": "Store Execution",
            "recommended_action": f"Investigate store {r['store_id']} at {r['retailer']}. SPI is {r['spi']:.1f} with {format_currency(r['revenue_opportunity'])} opportunity."
        })
    if len(distribution_gap):
        r = distribution_gap.sort_values("distribution_gap_count", ascending=False).iloc[0]
        recommendations.append({
            "priority": "High" if r["distribution_gap_count"] >= 10 else "Medium",
            "type": "Distribution Expansion",
            "recommended_action": f"Expand {r['brand']} in {r['retailer']}. Gap is {int(r['distribution_gap_count'])} stores."
        })
    if len(momentum):
        up = momentum[momentum["momentum_flag"] == "Trending Up"]
        if len(up):
            r = up.sort_values("momentum_ratio", ascending=False).iloc[0]
            recommendations.append({
                "priority": "Medium",
                "type": "Momentum Winner",
                "recommended_action": f"Increase support behind SKU {r['sku_id']} ({r['brand']}). Momentum ratio is {r['momentum_ratio']:.2f}."
            })
    if len(shelf_metrics):
        winners = shelf_metrics[shelf_metrics["shelf_action"] == "Increase Facings"]
        if len(winners):
            r = winners.sort_values("space_efficiency_index", ascending=False).iloc[0]
            recommendations.append({
                "priority": "Medium",
                "type": "Shelf Expansion",
                "recommended_action": f"Increase facings for SKU {r['sku_id']} ({r['brand']}). SEI is {r['space_efficiency_index']:.1f}."
            })

    sellin = []
    if len(distribution_gap):
        for _, r in distribution_gap.sort_values("distribution_gap_count", ascending=False).head(6).iterrows():
            sellin.append({
                "priority": "High" if r["distribution_gap_count"] >= 10 else "Medium",
                "retailer": r["retailer"],
                "sku_or_brand": r["brand"],
                "action": "Expand distribution",
                "rationale": f"Distribution gap of {int(r['distribution_gap_count'])} stores in {r['retailer']}."
            })

    return {
        "summary": summary,
        "store_perf": store_perf,
        "underperf": underperf,
        "distribution_gap": distribution_gap,
        "sku_velocity": sku_velocity,
        "yoy": yoy,
        "momentum": momentum,
        "shelf_metrics": shelf_metrics,
        "recommendations": pd.DataFrame(recommendations),
        "sellin": pd.DataFrame(sellin),
        "data_quality": pd.DataFrame([{
            "data_quality_score": data_quality_score,
            "invalid_dates": invalid_dates,
            "unmatched_skus": unmatched_skus,
            "unmatched_stores": unmatched_stores,
            "negative_units_pct": round(negative_units_pct, 2),
            "return_impact_pct": return_impact_pct
        }])
    }

st.title("ShelfIQ 911")
st.caption("Retail Analytics • Distribution Intelligence • Shelf Optimization")

with st.sidebar:
    st.markdown('<div class="sidebar-title">Navigation</div>', unsafe_allow_html=True)
    page = st.radio("Sections", ["Executive Dashboard", "Analytics Detail"], label_visibility="collapsed")
    st.markdown("---")
    mode = st.radio("Data mode", ["Demo data", "Client upload"])
    st.markdown("### Required Excel sheets")
    st.markdown("""
- `Sales_History`
- `Products`
- `Stores`
- `Shelf_Snapshot` *(optional)*
""")

sales = products = stores = shelf = None

if mode == "Client upload":
    uploaded_file = st.file_uploader("Upload Client Excel File", type=["xlsx", "xls"])
    if uploaded_file is not None:
        try:
            sales = safe_read_excel(uploaded_file, "Sales_History")
            products = safe_read_excel(uploaded_file, "Products")
            stores = safe_read_excel(uploaded_file, "Stores")
            try:
                shelf = safe_read_excel(uploaded_file, "Shelf_Snapshot")
            except Exception:
                shelf = pd.DataFrame()
            st.success("Client workbook loaded successfully.")
            with st.expander("Preview detected columns"):
                st.write("Sales_History:", list(sales.columns))
                st.write("Products:", list(products.columns))
                st.write("Stores:", list(stores.columns))
                if shelf is not None and len(shelf) > 0:
                    st.write("Shelf_Snapshot:", list(shelf.columns))
        except Exception as e:
            st.error(f"Could not read workbook: {e}")
else:
    sales, products, stores, shelf = generate_demo_data()
    st.info("Running in demo mode with sample data.")

run_clicked = st.button("Run ShelfIQ 911 Analysis", type="primary", use_container_width=True)

if run_clicked:
    if sales is None or products is None or stores is None:
        st.error("Please load data first.")
        st.stop()

    try:
        results = run_engine(sales, products, stores, shelf)
        summary = results["summary"]
        story = build_story(summary)

        if page == "Executive Dashboard":
            k1, k2, k3, k4, k5 = st.columns(5)
            with k1:
                st.markdown(metric_card_html("Retail Health", f"{summary['retail_health_score']}", summary["health_label"], "blue"), unsafe_allow_html=True)
            with k2:
                st.markdown(metric_card_html("Data Quality", f"{summary['data_quality_score']}", "Quality score", "purple"), unsafe_allow_html=True)
            with k3:
                st.markdown(metric_card_html("Revenue Opportunity", format_currency(summary["revenue_opportunity"]), "Modeled upside", "teal"), unsafe_allow_html=True)
            with k4:
                st.markdown(metric_card_html("Distribution Gap", f"{summary['avg_distribution_gap_pct']:.1f}%", "Average gap", "orange"), unsafe_allow_html=True)
            with k5:
                st.markdown(metric_card_html("Return Impact", f"{summary['return_impact_pct']:.1f}%", "Negative units impact", "dark"), unsafe_allow_html=True)

            st.markdown("---")
            st.markdown("## Executive Dashboard")
            top_left, top_right = st.columns([1.2, 1])

            with top_left:
                st.markdown(
                    f'''
                    <div class="story-box">
                        <div class="story-head">McKinsey Auto Story Generator</div>
                        <div class="story-body"><b>Headline:</b> {story["headline"]}</div>
                        <div class="story-body" style="margin-top:10px;"><b>So what:</b> {story["so_what"]}</div>
                        <div class="story-body" style="margin-top:10px;"><b>Key risk:</b> {story["risk"]}</div>
                    </div>
                    ''',
                    unsafe_allow_html=True
                )

            with top_right:
                st.markdown('<div class="panel">', unsafe_allow_html=True)
                st.markdown('<div style="font-size:1.1rem;font-weight:800;margin-bottom:8px;color:#0f172a;">Executive Scorecards</div>', unsafe_allow_html=True)
                g1, g2 = st.columns(2)
                with g1:
                    fig_g1 = go.Figure(go.Indicator(
                        mode="gauge+number",
                        value=summary["retail_health_score"],
                        title={"text":"Retail Health"},
                        gauge={"axis":{"range":[0,100]}, "bar":{"color":"#2563eb"}}
                    ))
                    fig_g1.update_layout(height=250, margin=dict(l=10,r=10,t=50,b=10))
                    st.plotly_chart(fig_g1, use_container_width=True)
                with g2:
                    fig_g2 = go.Figure(go.Indicator(
                        mode="gauge+number",
                        value=summary["data_quality_score"],
                        title={"text":"Data Quality"},
                        gauge={"axis":{"range":[0,100]}, "bar":{"color":"#7c3aed"}}
                    ))
                    fig_g2.update_layout(height=250, margin=dict(l=10,r=10,t=50,b=10))
                    st.plotly_chart(fig_g2, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

            st.markdown("### Why it matters")
            cards = story["evidence"]
            while len(cards) < 4:
                cards.append("More complete retailer data will unlock additional executive evidence.")
            e1, e2, e3, e4 = st.columns(4)
            for i, col in enumerate([e1, e2, e3, e4]):
                with col:
                    st.markdown(
                        f'''
                        <div class="story-card">
                            <div class="story-card-title">Evidence {i+1}</div>
                            <div class="story-card-body">{cards[i]}</div>
                        </div>
                        ''',
                        unsafe_allow_html=True
                    )

            st.markdown("### Executive Visuals")
            c1, c2, c3 = st.columns(3)
            with c1:
                temp = results["sku_velocity"].sort_values("sku_velocity_index", ascending=False).head(10).copy()
                fig = px.bar(temp, x="sku_id", y="sku_velocity_index", title="SKU Velocity")
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                temp = results["store_perf"].sort_values("revenue_opportunity", ascending=False).head(10).copy()
                fig = px.bar(temp, x="store_id", y="revenue_opportunity", title="Store Performance")
                st.plotly_chart(fig, use_container_width=True)
            with c3:
                temp = results["distribution_gap"].sort_values("distribution_gap_count", ascending=False).head(10).copy()
                fig = px.bar(temp, x="brand", y="distribution_gap_count", title="Distribution Gaps")
                st.plotly_chart(fig, use_container_width=True)

            st.markdown("## Strategic Action Center")
            a1, a2 = st.columns(2)
            with a1:
                recs = results["recommendations"]
                if len(recs):
                    for _, row in recs.iterrows():
                        klass = "medium"
                        if str(row["priority"]).lower() == "high":
                            klass = "high"
                        elif str(row["priority"]).lower() == "low":
                            klass = "low"
                        st.markdown(
                            f'<div class="action-box {klass}"><b>{row["type"]}</b><br>{row["recommended_action"]}</div>',
                            unsafe_allow_html=True
                        )
                else:
                    st.info("No recommendation actions available.")
            with a2:
                sellin = results["sellin"]
                if len(sellin):
                    for _, row in sellin.iterrows():
                        klass = "medium"
                        if str(row["priority"]).lower() == "high":
                            klass = "high"
                        elif str(row["priority"]).lower() == "low":
                            klass = "low"
                        st.markdown(
                            f'<div class="action-box {klass}"><b>{row["action"]}</b><br>{row["retailer"]} • {row["sku_or_brand"]}<br>{row["rationale"]}</div>',
                            unsafe_allow_html=True
                        )
                else:
                    st.info("No sell-in actions available.")
        else:
            st.markdown("## Analytics Detail")
            tabs = st.tabs([
                "Store Performance", "SKU Velocity", "Distribution",
                "YoY Growth", "Momentum", "Shelf Productivity", "Data Quality"
            ])
            with tabs[0]:
                st.dataframe(results["store_perf"], use_container_width=True)
            with tabs[1]:
                st.dataframe(results["sku_velocity"], use_container_width=True)
            with tabs[2]:
                st.dataframe(results["distribution_gap"], use_container_width=True)
            with tabs[3]:
                if len(results["yoy"]):
                    st.dataframe(results["yoy"], use_container_width=True)
                else:
                    st.info("Not enough year history for YoY output.")
            with tabs[4]:
                st.dataframe(results["momentum"], use_container_width=True)
            with tabs[5]:
                if len(results["shelf_metrics"]):
                    st.dataframe(results["shelf_metrics"], use_container_width=True)
                else:
                    st.info("No shelf snapshot uploaded.")
            with tabs[6]:
                st.dataframe(results["data_quality"], use_container_width=True)

    except Exception as e:
        st.error(f"Analysis failed: {e}")
