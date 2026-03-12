from io import BytesIO
import pandas as pd
import numpy as np
import streamlit as st
import matplotlib.pyplot as plt

st.set_page_config(page_title="ShelfIQ 911", layout="wide")

# =========================================================
# STATE -> REGION MAP
# =========================================================
STATE_TO_REGION = {
    "CT": "Northeast", "ME": "Northeast", "MA": "Northeast", "NH": "Northeast",
    "RI": "Northeast", "VT": "Northeast", "NJ": "Northeast", "NY": "Northeast",
    "PA": "Northeast",

    "IL": "Midwest", "IN": "Midwest", "MI": "Midwest", "OH": "Midwest",
    "WI": "Midwest", "IA": "Midwest", "KS": "Midwest", "MN": "Midwest",
    "MO": "Midwest", "NE": "Midwest", "ND": "Midwest", "SD": "Midwest",

    "AL": "Southeast", "AR": "Southeast", "DE": "Southeast", "DC": "Southeast",
    "FL": "Southeast", "GA": "Southeast", "KY": "Southeast", "LA": "Southeast",
    "MD": "Southeast", "MS": "Southeast", "NC": "Southeast", "SC": "Southeast",
    "TN": "Southeast", "VA": "Southeast", "WV": "Southeast",

    "AZ": "Southwest", "NM": "Southwest", "OK": "Southwest", "TX": "Southwest",

    "AK": "West", "CA": "West", "CO": "West", "HI": "West", "ID": "West",
    "MT": "West", "NV": "West", "OR": "West", "UT": "West", "WA": "West",
    "WY": "West"
}

# =========================================================
# HELPERS
# =========================================================
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]
    return df

def normalize_state(x):
    if pd.isna(x):
        return np.nan
    return str(x).strip().upper()

def read_excel_sheet(uploaded_file, sheet_name: str) -> pd.DataFrame:
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=0)
    return normalize_columns(df)

def read_uploaded_table(uploaded_file):
    if uploaded_file is None:
        return None

    name = uploaded_file.name.lower()

    if name.endswith(".csv"):
        return normalize_columns(pd.read_csv(uploaded_file))

    if name.endswith(".xlsx") or name.endswith(".xls"):
        uploaded_file.seek(0)
        return normalize_columns(pd.read_excel(uploaded_file, header=0))

    raise ValueError(f"Unsupported file type: {uploaded_file.name}")

def validate_inputs(products, stores, sales_history, shelf=None):
    required_product_cols = {"sku_id"}
    required_store_cols = {"store_id", "retailer", "state"}
    required_sales_cols = {"store_id", "sku_id", "week_end_date", "units"}

    missing = {}

    if not required_product_cols.issubset(products.columns):
        missing["Products"] = sorted(list(required_product_cols - set(products.columns)))

    if not required_store_cols.issubset(stores.columns):
        missing["Stores"] = sorted(list(required_store_cols - set(stores.columns)))

    if not required_sales_cols.issubset(sales_history.columns):
        missing["Sales_History"] = sorted(list(required_sales_cols - set(sales_history.columns)))

    if shelf is not None and len(shelf) > 0:
        required_shelf_cols = {"store_id", "sku_id"}
        if not required_shelf_cols.issubset(shelf.columns):
            missing["Shelf_Snapshot"] = sorted(list(required_shelf_cols - set(shelf.columns)))

    return missing

def prepare_stores(stores: pd.DataFrame) -> pd.DataFrame:
    stores = normalize_columns(stores).copy()

    stores["store_id"] = stores["store_id"].astype(str).str.strip()
    stores["retailer"] = stores["retailer"].astype(str).str.strip()

    if "format" not in stores.columns:
        stores["format"] = "Unknown"
    else:
        stores["format"] = stores["format"].fillna("Unknown").astype(str).str.strip()

    if "store_name" not in stores.columns:
        stores["store_name"] = np.nan

    stores["state"] = stores["state"].apply(normalize_state)
    stores["region"] = stores["state"].map(STATE_TO_REGION)

    return stores

def classify_quality_score(score: float) -> str:
    if score >= 95:
        return "Excellent"
    if score >= 85:
        return "Good"
    if score >= 70:
        return "Fair"
    return "Needs Cleanup"

def classify_health_score(score: float) -> str:
    if score >= 90:
        return "Excellent"
    if score >= 75:
        return "Strong"
    if score >= 60:
        return "Fair"
    return "Weak"

def run_data_quality_checks(products, stores, sales_history, shelf=None):
    issues = []

    products = normalize_columns(products).copy()
    stores = normalize_columns(stores).copy()
    sales_history = normalize_columns(sales_history).copy()

    if shelf is not None and len(shelf) > 0:
        shelf = normalize_columns(shelf).copy()
    else:
        shelf = pd.DataFrame()

    # Keep original row count
    rows_uploaded = len(sales_history)

    if "store_id" in stores.columns:
        stores["store_id"] = stores["store_id"].astype(str).str.strip()
    if "sku_id" in products.columns:
        products["sku_id"] = products["sku_id"].astype(str).str.strip()

    if "store_id" in sales_history.columns:
        sales_history["store_id"] = sales_history["store_id"].astype(str).str.strip()
    if "sku_id" in sales_history.columns:
        sales_history["sku_id"] = sales_history["sku_id"].astype(str).str.strip()

    if "state" in stores.columns:
        stores["state"] = stores["state"].apply(normalize_state)

    def add_issue(check, status, count, severity_weight):
        issues.append({
            "check": check,
            "status": status,
            "count": int(count),
            "severity_weight": severity_weight
        })

    # Missing IDs
    if "store_id" in sales_history.columns:
        missing_store_ids = sales_history["store_id"].isna().sum() + (sales_history["store_id"].astype(str).str.strip() == "").sum()
        add_issue("Missing store_id in Sales_History", "Fail" if missing_store_ids > 0 else "Pass", missing_store_ids, 8)

    if "sku_id" in sales_history.columns:
        missing_sku_ids = sales_history["sku_id"].isna().sum() + (sales_history["sku_id"].astype(str).str.strip() == "").sum()
        add_issue("Missing sku_id in Sales_History", "Fail" if missing_sku_ids > 0 else "Pass", missing_sku_ids, 8)

    # Dates
    invalid_dates = 0
    if "week_end_date" in sales_history.columns:
        parsed_dates = pd.to_datetime(sales_history["week_end_date"], errors="coerce")
        invalid_dates = parsed_dates.isna().sum()
        add_issue("Invalid week_end_date values", "Fail" if invalid_dates > 0 else "Pass", invalid_dates, 7)
        sales_history["week_end_date"] = parsed_dates

    # Units
    negative_units = 0
    negative_units_pct = 0.0
    if "units" in sales_history.columns:
        sales_history["units"] = pd.to_numeric(sales_history["units"], errors="coerce")
        non_numeric_units = sales_history["units"].isna().sum()
        add_issue("Non-numeric units", "Fail" if non_numeric_units > 0 else "Pass", non_numeric_units, 7)
        sales_history["units"] = sales_history["units"].fillna(0)
        negative_units = (sales_history["units"] < 0).sum()
        negative_units_pct = round((negative_units / max(len(sales_history), 1)) * 100, 2)
        neg_units_status = "Pass"
        if negative_units_pct > 20:
            neg_units_status = "Fail"
        elif negative_units_pct > 5:
            neg_units_status = "Warn"
        add_issue("Negative units present", neg_units_status, negative_units, 4)

    # Sales
    negative_sales = 0
    negative_sales_pct = 0.0
    sales_col = None
    if "sales_dollars" in sales_history.columns:
        sales_col = "sales_dollars"
    elif "sales" in sales_history.columns:
        sales_col = "sales"

    if sales_col is not None:
        sales_history[sales_col] = pd.to_numeric(sales_history[sales_col], errors="coerce")
        non_numeric_sales = sales_history[sales_col].isna().sum()
        add_issue(f"Non-numeric {sales_col}", "Fail" if non_numeric_sales > 0 else "Pass", non_numeric_sales, 7)
        sales_history[sales_col] = sales_history[sales_col].fillna(0)
        negative_sales = (sales_history[sales_col] < 0).sum()
        negative_sales_pct = round((negative_sales / max(len(sales_history), 1)) * 100, 2)
        neg_sales_status = "Pass"
        if negative_sales_pct > 20:
            neg_sales_status = "Fail"
        elif negative_sales_pct > 5:
            neg_sales_status = "Warn"
        add_issue(f"Negative {sales_col} present", neg_sales_status, negative_sales, 4)

    # Duplicates
    dup_cols = [c for c in ["store_id", "sku_id", "week_end_date"] if c in sales_history.columns]
    dup_count = 0
    if len(dup_cols) == 3:
        dup_count = sales_history.duplicated(subset=dup_cols).sum()
        add_issue("Duplicate store_id + sku_id + week_end_date rows", "Warn" if dup_count > 0 else "Pass", dup_count, 3)

    # Unmatcheds
    unmatched_skus_count = 0
    unmatched_stores_count = 0
    if "sku_id" in sales_history.columns and "sku_id" in products.columns:
        unmatched_skus = ~sales_history["sku_id"].isin(products["sku_id"])
        unmatched_skus_count = int(unmatched_skus.sum())
        add_issue("Sales_History sku_id not found in Products", "Fail" if unmatched_skus_count > 0 else "Pass", unmatched_skus_count, 6)

    if "store_id" in sales_history.columns and "store_id" in stores.columns:
        unmatched_stores = ~sales_history["store_id"].isin(stores["store_id"])
        unmatched_stores_count = int(unmatched_stores.sum())
        add_issue("Sales_History store_id not found in Stores", "Fail" if unmatched_stores_count > 0 else "Pass", unmatched_stores_count, 6)

    # States
    invalid_states_count = 0
    unmapped_region_count = 0
    if "state" in stores.columns:
        invalid_states = ~stores["state"].isin(list(STATE_TO_REGION.keys()))
        invalid_states_count = int(invalid_states.sum())
        add_issue("Invalid state codes in Stores", "Fail" if invalid_states_count > 0 else "Pass", invalid_states_count, 5)

        missing_region_map = stores["state"].map(STATE_TO_REGION).isna().sum()
        unmapped_region_count = int(missing_region_map)
        add_issue("States that could not be mapped to region", "Fail" if unmapped_region_count > 0 else "Pass", unmapped_region_count, 5)

    # Sparse week coverage
    sparse_pairs = 0
    if {"store_id", "sku_id", "week_end_date"}.issubset(sales_history.columns):
        counts = sales_history.groupby(["store_id", "sku_id"])["week_end_date"].nunique().reset_index(name="week_count")
        sparse_pairs = int((counts["week_count"] < counts["week_count"].median()).sum()) if len(counts) else 0
        add_issue("Store/SKU pairs with below-median week coverage", "Warn" if sparse_pairs > 0 else "Pass", sparse_pairs, 2)

    # Shelf
    bad_shelf_share = 0
    bad_facings = 0
    if len(shelf) > 0:
        if "shelf_share" in shelf.columns:
            shelf["shelf_share"] = pd.to_numeric(shelf["shelf_share"], errors="coerce")
            bad_shelf_share = int(((shelf["shelf_share"] < 0) | (shelf["shelf_share"] > 1)).sum())
            add_issue("Shelf_Snapshot shelf_share outside 0 to 1", "Fail" if bad_shelf_share > 0 else "Pass", bad_shelf_share, 5)

        if "facings" in shelf.columns:
            shelf["facings"] = pd.to_numeric(shelf["facings"], errors="coerce")
            bad_facings = int((shelf["facings"] < 0).sum())
            add_issue("Negative facings in Shelf_Snapshot", "Fail" if bad_facings > 0 else "Pass", bad_facings, 5)

    quality = pd.DataFrame(issues)

    # Score
    penalty = 0
    for _, row in quality.iterrows():
        if row["status"] == "Fail":
            penalty += min(row["count"], 10) * row["severity_weight"]
        elif row["status"] == "Warn":
            penalty += min(row["count"], 10) * row["severity_weight"] * 0.35

    quality_score = max(0, round(100 - penalty, 1))

    # Accepted vs rejected rows (strictest rows rejected)
    rejected_mask = pd.Series(False, index=sales_history.index)

    if "store_id" in sales_history.columns:
        rejected_mask = rejected_mask | sales_history["store_id"].isna() | (sales_history["store_id"].astype(str).str.strip() == "")
    if "sku_id" in sales_history.columns:
        rejected_mask = rejected_mask | sales_history["sku_id"].isna() | (sales_history["sku_id"].astype(str).str.strip() == "")
    if "week_end_date" in sales_history.columns:
        rejected_mask = rejected_mask | sales_history["week_end_date"].isna()

    if "units" in sales_history.columns:
        rejected_mask = rejected_mask | sales_history["units"].isna()

    rows_rejected = int(rejected_mask.sum())
    rows_accepted = int(rows_uploaded - rows_rejected)

    # Return impact
    abs_total_units = float(sales_history["units"].abs().sum()) if "units" in sales_history.columns else 0
    abs_negative_units = float(sales_history.loc[sales_history["units"] < 0, "units"].abs().sum()) if "units" in sales_history.columns else 0
    return_impact_score = round((abs_negative_units / abs_total_units) * 100, 2) if abs_total_units > 0 else 0.0

    meta = {
        "rows_uploaded": rows_uploaded,
        "rows_accepted": rows_accepted,
        "rows_rejected": rows_rejected,
        "negative_units_pct": negative_units_pct,
        "negative_sales_pct": negative_sales_pct,
        "return_impact_score": return_impact_score,
        "quality_score": quality_score,
        "quality_label": classify_quality_score(quality_score),
    }

    return quality, meta

def build_recommendations(underperf, dist, yoy, momentum, shelf_efficiency):
    recommendations = []

    if len(underperf):
        row = underperf.sort_values("revenue_opportunity_score", ascending=False).iloc[0]
        recommendations.append(
            f"Investigate store {row['store_id']} at {row['retailer']} in {row['region']}: "
            f"revenue opportunity is ${row['revenue_opportunity_score']:,.0f} and SPI is {row['store_performance_index']:.1f}."
        )

    if len(dist):
        row = dist.sort_values("distribution_gap_count", ascending=False).iloc[0]
        recommendations.append(
            f"Expand distribution for {row['brand']} / {row['category']} at {row['retailer']}: "
            f"gap of {int(row['distribution_gap_count'])} stores."
        )

    if len(yoy):
        yoy_clean = yoy.dropna(subset=["yoy_sales_growth_pct"])
        if len(yoy_clean):
            top = yoy_clean.sort_values("yoy_sales_growth_pct", ascending=False).iloc[0]
            recommendations.append(
                f"Protect and expand SKU {top['sku_id']} ({top['brand']}): "
                f"YoY sales growth is {top['yoy_sales_growth_pct']:.1f}%."
            )

            bottom = yoy_clean.sort_values("yoy_sales_growth_pct", ascending=True).iloc[0]
            recommendations.append(
                f"Review SKU {bottom['sku_id']} ({bottom['brand']}): "
                f"YoY sales growth is {bottom['yoy_sales_growth_pct']:.1f}%."
            )

    if len(momentum):
        up = momentum[momentum["momentum_flag"] == "Trending Up"]
        down = momentum[momentum["momentum_flag"] == "Trending Down"]

        if len(up):
            row = up.sort_values("momentum_ratio", ascending=False).iloc[0]
            recommendations.append(
                f"Increase support behind momentum winner {row['sku_id']} ({row['brand']}): "
                f"momentum ratio is {row['momentum_ratio']:.2f}."
            )

        if len(down):
            row = down.sort_values("momentum_ratio", ascending=True).iloc[0]
            recommendations.append(
                f"Diagnose decline for {row['sku_id']} ({row['brand']}): "
                f"momentum ratio is {row['momentum_ratio']:.2f}."
            )

    if len(shelf_efficiency):
        winners = shelf_efficiency[shelf_efficiency["shelf_action"] == "Increase Facings"]
        losers = shelf_efficiency[shelf_efficiency["shelf_action"] == "Reduce / Review"]

        if len(winners):
            row = winners.sort_values("space_efficiency_index", ascending=False).iloc[0]
            recommendations.append(
                f"Increase facings for SKU {row['sku_id']} ({row['brand']}): "
                f"Space Efficiency Index is {row['space_efficiency_index']:.1f}."
            )

        if len(losers):
            row = losers.sort_values("space_efficiency_index", ascending=True).iloc[0]
            recommendations.append(
                f"Review shelf space for SKU {row['sku_id']} ({row['brand']}): "
                f"Space Efficiency Index is {row['space_efficiency_index']:.1f}."
            )

    return pd.DataFrame({"recommended_action": recommendations[:8]})

def build_sell_in_engine(dist, momentum, yoy, shelf_efficiency, underperf):
    sell_in_rows = []

    if len(dist):
        dist_top = dist.sort_values(["distribution_gap_count", "distribution_gap_index"], ascending=[False, False]).head(10)
        for _, row in dist_top.iterrows():
            sell_in_rows.append({
                "priority": "High" if row["distribution_gap_count"] >= 10 else "Medium",
                "retailer": row["retailer"],
                "sku_or_brand": row["brand"],
                "action": "Expand distribution",
                "rationale": f"Distribution gap of {int(row['distribution_gap_count'])} stores in {row['retailer']}.",
                "estimated_opportunity": np.nan
            })

    if len(shelf_efficiency):
        winners = shelf_efficiency[shelf_efficiency["shelf_action"] == "Increase Facings"].sort_values(
            "space_efficiency_index", ascending=False
        ).head(10)
        for _, row in winners.iterrows():
            sell_in_rows.append({
                "priority": "High" if row["space_efficiency_index"] >= 140 else "Medium",
                "retailer": row.get("retailer", "Mixed"),
                "sku_or_brand": row["sku_id"],
                "action": "Increase facings",
                "rationale": f"Space Efficiency Index of {row['space_efficiency_index']:.1f} with strong shelf productivity.",
                "estimated_opportunity": row.get("total_sales", np.nan)
            })

    if len(underperf):
        top_exec = underperf.sort_values("revenue_opportunity_score", ascending=False).head(10)
        for _, row in top_exec.iterrows():
            sell_in_rows.append({
                "priority": row["opportunity_priority"],
                "retailer": row["retailer"],
                "sku_or_brand": row["store_id"],
                "action": "Fix store execution",
                "rationale": f"SPI of {row['store_performance_index']:.1f} and revenue opportunity of ${row['revenue_opportunity_score']:,.0f}.",
                "estimated_opportunity": row["revenue_opportunity_score"]
            })

    if len(momentum):
        movers = momentum[momentum["momentum_flag"] == "Trending Up"].sort_values("momentum_ratio", ascending=False).head(10)
        for _, row in movers.iterrows():
            sell_in_rows.append({
                "priority": "Medium",
                "retailer": "Mixed",
                "sku_or_brand": row["sku_id"],
                "action": "Sell-in support",
                "rationale": f"Momentum ratio of {row['momentum_ratio']:.2f} indicates strong recent acceleration.",
                "estimated_opportunity": np.nan
            })

    sell_in = pd.DataFrame(sell_in_rows)
    if len(sell_in):
        sell_in = sell_in.drop_duplicates(subset=["retailer", "sku_or_brand", "action"]).reset_index(drop=True)
    return sell_in

def run_analysis(products, stores, sales_history, shelf=None):
    products = normalize_columns(products).copy()
    sales_history = normalize_columns(sales_history).copy()

    products["sku_id"] = products["sku_id"].astype(str).str.strip()

    if "brand" not in products.columns:
        products["brand"] = "Unknown"
    else:
        products["brand"] = products["brand"].fillna("Unknown")

    if "category" not in products.columns:
        products["category"] = "Unknown"
    else:
        products["category"] = products["category"].fillna("Unknown")

    stores = prepare_stores(stores)

    sales_history["store_id"] = sales_history["store_id"].astype(str).str.strip()
    sales_history["sku_id"] = sales_history["sku_id"].astype(str).str.strip()
    sales_history["week_end_date"] = pd.to_datetime(sales_history["week_end_date"], errors="coerce")
    sales_history["units"] = pd.to_numeric(sales_history["units"], errors="coerce").fillna(0)

    if "sales_dollars" in sales_history.columns:
        sales_history["sales_dollars"] = pd.to_numeric(sales_history["sales_dollars"], errors="coerce").fillna(0)
    elif "sales" in sales_history.columns:
        sales_history["sales_dollars"] = pd.to_numeric(sales_history["sales"], errors="coerce").fillna(0)
    else:
        sales_history["sales_dollars"] = 0

    if shelf is None:
        shelf = pd.DataFrame(columns=["store_id", "sku_id", "facings", "shelf_share"])
    else:
        shelf = normalize_columns(shelf).copy()
        if "store_id" in shelf.columns:
            shelf["store_id"] = shelf["store_id"].astype(str).str.strip()
        if "sku_id" in shelf.columns:
            shelf["sku_id"] = shelf["sku_id"].astype(str).str.strip()
        if "facings" in shelf.columns:
            shelf["facings"] = pd.to_numeric(shelf["facings"], errors="coerce").fillna(0)
        if "shelf_share" in shelf.columns:
            shelf["shelf_share"] = pd.to_numeric(shelf["shelf_share"], errors="coerce").fillna(0)

    quality_checks, quality_meta = run_data_quality_checks(products, stores, sales_history, shelf)

    sales_enriched = (
        sales_history
        .merge(products, on="sku_id", how="left")
        .merge(stores, on="store_id", how="left")
    )

    for col in ["brand", "category", "retailer", "format", "state", "region"]:
        if col not in sales_enriched.columns:
            sales_enriched[col] = "Unknown"
        else:
            sales_enriched[col] = sales_enriched[col].fillna("Unknown")

    current_max_week = sales_enriched["week_end_date"].max()
    if pd.isna(current_max_week):
        raise ValueError("No valid dates found in Sales_History.")

    trailing_13w_start = current_max_week - pd.Timedelta(weeks=13)
    trailing_52w_start = current_max_week - pd.Timedelta(weeks=52)

    sales_13w = sales_enriched[sales_enriched["week_end_date"] > trailing_13w_start].copy()
    sales_52w = sales_enriched[sales_enriched["week_end_date"] > trailing_52w_start].copy()

    weeks_13 = max(sales_13w["week_end_date"].nunique(), 1)
    weeks_52 = max(sales_52w["week_end_date"].nunique(), 1)

    # SKU Velocity
    sku_velocity = (
        sales_13w
        .groupby(["sku_id", "brand", "category"], dropna=False)
        .agg(
            total_units=("units", "sum"),
            total_sales=("sales_dollars", "sum"),
            active_stores=("store_id", "nunique")
        )
        .reset_index()
    )

    sku_velocity["velocity_units_per_store_per_week"] = (
        sku_velocity["total_units"] /
        sku_velocity["active_stores"].clip(lower=1) /
        weeks_13
    )

    category_avg_velocity = (
        sku_velocity.groupby("category", dropna=False)["velocity_units_per_store_per_week"]
        .mean()
        .rename("category_avg_velocity")
        .reset_index()
    )

    sku_velocity = sku_velocity.merge(category_avg_velocity, on="category", how="left")
    sku_velocity["sku_velocity_index"] = (
        sku_velocity["velocity_units_per_store_per_week"] /
        sku_velocity["category_avg_velocity"].replace(0, np.nan)
    ) * 100
    sku_velocity["sku_velocity_index"] = sku_velocity["sku_velocity_index"].fillna(0)

    # Store Performance
    store_totals = (
        sales_13w
        .groupby(["store_id", "retailer", "region", "state", "format"], dropna=False)
        .agg(
            actual_sales=("sales_dollars", "sum"),
            actual_units=("units", "sum"),
            sku_count=("sku_id", "nunique")
        )
        .reset_index()
    )

    peer_avg = (
        store_totals
        .groupby(["retailer", "format", "region"], dropna=False)["actual_sales"]
        .mean()
        .rename("expected_sales")
        .reset_index()
    )

    store_perf = store_totals.merge(peer_avg, on=["retailer", "format", "region"], how="left")
    store_perf["expected_sales"] = store_perf["expected_sales"].fillna(store_perf["actual_sales"].mean())
    store_perf["store_performance_index"] = (
        store_perf["actual_sales"] /
        store_perf["expected_sales"].replace(0, np.nan)
    ) * 100
    store_perf["store_performance_index"] = store_perf["store_performance_index"].fillna(0)
    store_perf["sales_gap"] = store_perf["expected_sales"] - store_perf["actual_sales"]
    store_perf["underperforming_flag"] = store_perf["store_performance_index"] < 80

    retailer_avg_spi = (
        store_perf.groupby("retailer", dropna=False)["store_performance_index"]
        .mean()
        .rename("retailer_avg_spi")
        .reset_index()
    )
    store_perf = store_perf.merge(retailer_avg_spi, on="retailer", how="left")

    store_perf["opportunity_confidence"] = np.select(
        [
            (store_perf["store_performance_index"] < 70) & (quality_meta["quality_score"] >= 85),
            (store_perf["store_performance_index"] < 85) & (quality_meta["quality_score"] >= 70),
        ],
        ["High", "Medium"],
        default="Low"
    )

    confidence_factor = np.select(
        [
            store_perf["opportunity_confidence"] == "High",
            store_perf["opportunity_confidence"] == "Medium",
            store_perf["opportunity_confidence"] == "Low",
        ],
        [1.0, 0.7, 0.4],
        default=0.4
    )

    store_perf["revenue_opportunity_score"] = np.where(
        (store_perf["sales_gap"] > 0) & (store_perf["sales_gap"] >= 50),
        store_perf["sales_gap"] * confidence_factor,
        0
    )

    store_perf["opportunity_priority"] = np.select(
        [
            store_perf["revenue_opportunity_score"] >= 500,
            (store_perf["revenue_opportunity_score"] >= 200) & (store_perf["revenue_opportunity_score"] < 500),
            (store_perf["revenue_opportunity_score"] > 0) & (store_perf["revenue_opportunity_score"] < 200)
        ],
        ["High", "Medium", "Low"],
        default="None"
    )

    # Exception flags
    exception_flags = []
    for _, row in store_perf.iterrows():
        flags = []
        if row["store_performance_index"] < 70:
            flags.append("Store Execution Problem")
        if row["revenue_opportunity_score"] >= 500:
            flags.append("High Revenue Opportunity")
        exception_flags.append(" | ".join(flags) if flags else "Normal")
    store_perf["exception_flags"] = exception_flags

    underperforming_stores = (
        store_perf[store_perf["underperforming_flag"]]
        .sort_values(["revenue_opportunity_score", "store_performance_index"], ascending=[False, True])
        .reset_index(drop=True)
    )

    # Distribution Gap
    carried = (
        sales_13w
        .groupby(["brand", "category", "retailer", "store_id"], dropna=False)
        .agg(total_units=("units", "sum"))
        .reset_index()
    )
    carried = carried[carried["total_units"] > 0]

    retailer_store_universe = (
        stores
        .groupby("retailer", dropna=False)["store_id"]
        .nunique()
        .rename("retailer_store_universe")
        .reset_index()
    )

    brand_distribution = (
        carried
        .groupby(["brand", "category", "retailer"], dropna=False)["store_id"]
        .nunique()
        .rename("current_store_count")
        .reset_index()
        .merge(retailer_store_universe, on="retailer", how="left")
    )

    brand_distribution["distribution_gap_count"] = (
        brand_distribution["retailer_store_universe"] - brand_distribution["current_store_count"]
    ).clip(lower=0)

    brand_distribution["distribution_gap_index"] = (
        brand_distribution["distribution_gap_count"] /
        brand_distribution["retailer_store_universe"].replace(0, np.nan)
    ) * 100
    brand_distribution["distribution_gap_index"] = brand_distribution["distribution_gap_index"].fillna(0)

    # YoY
    sales_enriched["year"] = sales_enriched["week_end_date"].dt.year

    yearly_sku = (
        sales_enriched
        .groupby(["sku_id", "brand", "category", "year"], dropna=False)
        .agg(
            yearly_sales=("sales_dollars", "sum"),
            yearly_units=("units", "sum")
        )
        .reset_index()
    )

    years_available = sorted([y for y in yearly_sku["year"].dropna().unique()])
    yoy_growth = pd.DataFrame()

    if len(years_available) >= 2:
        prior_year = years_available[-2]
        current_year = years_available[-1]

        prior_df = yearly_sku[yearly_sku["year"] == prior_year][["sku_id", "yearly_sales", "yearly_units"]].rename(
            columns={"yearly_sales": f"sales_{prior_year}", "yearly_units": f"units_{prior_year}"}
        )

        current_df = yearly_sku[yearly_sku["year"] == current_year][["sku_id", "brand", "category", "yearly_sales", "yearly_units"]].rename(
            columns={"yearly_sales": f"sales_{current_year}", "yearly_units": f"units_{current_year}"}
        )

        yoy_growth = current_df.merge(prior_df, on="sku_id", how="left")

        yoy_growth[f"sales_{prior_year}"] = yoy_growth[f"sales_{prior_year}"].fillna(0)
        yoy_growth[f"units_{prior_year}"] = yoy_growth[f"units_{prior_year}"].fillna(0)

        yoy_growth["yoy_sales_growth_pct"] = np.where(
            yoy_growth[f"sales_{prior_year}"] > 0,
            ((yoy_growth[f"sales_{current_year}"] - yoy_growth[f"sales_{prior_year}"]) / yoy_growth[f"sales_{prior_year}"]) * 100,
            np.nan
        )

        yoy_growth["yoy_units_growth_pct"] = np.where(
            yoy_growth[f"units_{prior_year}"] > 0,
            ((yoy_growth[f"units_{current_year}"] - yoy_growth[f"units_{prior_year}"]) / yoy_growth[f"units_{prior_year}"]) * 100,
            np.nan
        )

        yoy_growth["exception_flags"] = np.select(
            [
                yoy_growth["yoy_sales_growth_pct"] >= 20,
                yoy_growth["yoy_sales_growth_pct"] <= -10
            ],
            [
                "YoY Winner",
                "YoY Decliner"
            ],
            default="Normal"
        )

    # Momentum
    velocity_13w = (
        sales_13w
        .groupby(["sku_id"], dropna=False)
        .agg(units_13w=("units", "sum"), stores_13w=("store_id", "nunique"))
        .reset_index()
    )
    velocity_13w["velocity_13w"] = (
        velocity_13w["units_13w"] / velocity_13w["stores_13w"].clip(lower=1) / weeks_13
    )

    velocity_52w = (
        sales_52w
        .groupby(["sku_id"], dropna=False)
        .agg(units_52w=("units", "sum"), stores_52w=("store_id", "nunique"))
        .reset_index()
    )
    velocity_52w["velocity_52w"] = (
        velocity_52w["units_52w"] / velocity_52w["stores_52w"].clip(lower=1) / weeks_52
    )

    momentum = (
        velocity_13w.merge(velocity_52w, on="sku_id", how="outer")
        .merge(products[["sku_id", "brand", "category"]], on="sku_id", how="left")
    )

    momentum["velocity_13w"] = momentum["velocity_13w"].fillna(0)
    momentum["velocity_52w"] = momentum["velocity_52w"].fillna(0)

    momentum["momentum_ratio"] = np.where(
        momentum["velocity_52w"] > 0,
        momentum["velocity_13w"] / momentum["velocity_52w"],
        np.nan
    )

    momentum["momentum_flag"] = np.select(
        [
            momentum["momentum_ratio"] >= 1.20,
            momentum["momentum_ratio"] <= 0.80
        ],
        [
            "Trending Up",
            "Trending Down"
        ],
        default="Stable"
    )

    # Recent declines
    sku_declines = (
        sales_enriched
        .groupby(["sku_id", "brand", "category", "week_end_date"], dropna=False)
        .agg(weekly_sales=("sales_dollars", "sum"))
        .reset_index()
        .sort_values(["sku_id", "week_end_date"])
    )

    sku_declines["prev_week_sales"] = sku_declines.groupby("sku_id")["weekly_sales"].shift(1)
    sku_declines["wow_change_pct"] = (
        (sku_declines["weekly_sales"] - sku_declines["prev_week_sales"]) /
        sku_declines["prev_week_sales"].replace(0, np.nan)
    ) * 100
    sku_declines["wow_change_pct"] = sku_declines["wow_change_pct"].fillna(0)

    recent_declines = (
        sku_declines[sku_declines["wow_change_pct"] <= -10]
        .sort_values("wow_change_pct")
        .reset_index(drop=True)
    )

    # Shelf metrics + SEI
    shelf_metrics = pd.DataFrame()
    if len(shelf) > 0 and {"store_id", "sku_id"}.issubset(shelf.columns):
        if "facings" not in shelf.columns:
            shelf["facings"] = np.nan
        if "shelf_share" not in shelf.columns:
            shelf["shelf_share"] = np.nan

        shelf_metrics = (
            shelf
            .merge(products, on="sku_id", how="left")
            .merge(
                sales_13w.groupby(["store_id", "sku_id"], dropna=False)
                .agg(
                    total_sales=("sales_dollars", "sum"),
                    total_units=("units", "sum")
                )
                .reset_index(),
                on=["store_id", "sku_id"],
                how="left"
            )
            .merge(stores[["store_id", "retailer", "state", "region", "format"]], on="store_id", how="left")
        )

        shelf_metrics["total_sales"] = shelf_metrics["total_sales"].fillna(0)
        shelf_metrics["total_units"] = shelf_metrics["total_units"].fillna(0)
        shelf_metrics["facings"] = pd.to_numeric(shelf_metrics["facings"], errors="coerce")
        shelf_metrics["shelf_productivity_score"] = (
            shelf_metrics["total_sales"] / shelf_metrics["facings"].replace(0, np.nan)
        )
        shelf_metrics["shelf_productivity_score"] = shelf_metrics["shelf_productivity_score"].fillna(0)
        shelf_metrics["sales_per_facing"] = shelf_metrics["shelf_productivity_score"]

        category_avg_spf = (
            shelf_metrics.groupby("category", dropna=False)["sales_per_facing"]
            .mean()
            .rename("category_avg_sales_per_facing")
            .reset_index()
        )
        shelf_metrics = shelf_metrics.merge(category_avg_spf, on="category", how="left")

        shelf_metrics["space_efficiency_index"] = (
            shelf_metrics["sales_per_facing"] /
            shelf_metrics["category_avg_sales_per_facing"].replace(0, np.nan)
        ) * 100
        shelf_metrics["space_efficiency_index"] = shelf_metrics["space_efficiency_index"].fillna(0)

        shelf_metrics["shelf_action"] = np.select(
            [
                shelf_metrics["space_efficiency_index"] >= 120,
                shelf_metrics["space_efficiency_index"] < 80
            ],
            [
                "Increase Facings",
                "Reduce / Review"
            ],
            default="Hold"
        )

        shelf_metrics["exception_flags"] = np.select(
            [
                shelf_metrics["space_efficiency_index"] >= 120,
                shelf_metrics["space_efficiency_index"] < 80
            ],
            [
                "Shelf Space Winner",
                "Shelf Space Risk"
            ],
            default="Normal"
        )

    recommendations = build_recommendations(underperforming_stores, brand_distribution, yoy_growth, momentum, shelf_metrics)
    sell_in = build_sell_in_engine(brand_distribution, momentum, yoy_growth, shelf_metrics, underperforming_stores)

    # Health summary
    underperf_rate = float(store_perf["underperforming_flag"].mean()) if len(store_perf) else 0
    avg_spi = float(store_perf["store_performance_index"].fillna(100).mean()) if len(store_perf) else 100
    dist_gap_rate = float(brand_distribution["distribution_gap_index"].fillna(0).mean()) if len(brand_distribution) else 0

    a = min(max(avg_spi, 0), 120) / 120 * 40
    b = (1 - min(max(underperf_rate, 0), 1)) * 15
    c = (1 - min(max(dist_gap_rate / 100, 0), 1)) * 15
    d = (quality_meta["quality_score"] / 100) * 20
    e = min(quality_meta["return_impact_score"], 10) / 10 * 10
    retail_health_score = round(a + b + c + d + e, 1)

    health_summary = pd.DataFrame([{
        "retail_health_score": retail_health_score,
        "retail_health_label": classify_health_score(retail_health_score),
        "data_quality_score": quality_meta["quality_score"],
        "data_quality_label": quality_meta["quality_label"],
        "rows_uploaded": quality_meta["rows_uploaded"],
        "rows_accepted": quality_meta["rows_accepted"],
        "rows_rejected": quality_meta["rows_rejected"],
        "return_impact_score": quality_meta["return_impact_score"],
        "negative_units_pct": quality_meta["negative_units_pct"],
        "negative_sales_pct": quality_meta["negative_sales_pct"],
        "store_count": int(store_perf["store_id"].nunique()),
        "sku_count": int(products["sku_id"].nunique()),
        "underperforming_store_count": int(store_perf["underperforming_flag"].sum()),
        "avg_store_performance_index": round(avg_spi, 2),
        "avg_distribution_gap_index": round(dist_gap_rate, 2),
        "estimated_revenue_opportunity": round(float(store_perf["revenue_opportunity_score"].sum()), 2),
    }])

    return {
        "health_summary": health_summary,
        "quality_checks": quality_checks,
        "recommendations": recommendations,
        "sell_in_opportunities": sell_in,
        "store_performance_index": store_perf,
        "underperforming_stores": underperforming_stores,
        "sku_velocity_score": sku_velocity,
        "distribution_gap_index": brand_distribution,
        "revenue_opportunity_score": store_perf,
        "recent_sku_declines": recent_declines,
        "shelf_productivity_score": shelf_metrics,
        "yoy_growth": yoy_growth,
        "momentum": momentum,
    }

def to_excel_download(results_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in results_dict.items():
            if isinstance(df, pd.DataFrame) and len(df) > 0:
                df.to_excel(writer, sheet_name=name[:31], index=False)
    output.seek(0)
    return output

def make_bar_chart(df, x_col, y_col, title, top_n=10, ascending=False):
    if df is None or len(df) == 0 or x_col not in df.columns or y_col not in df.columns:
        return None

    temp = df[[x_col, y_col]].dropna().copy()
    if len(temp) == 0:
        return None

    temp = temp.sort_values(y_col, ascending=ascending).head(top_n)
    fig, ax = plt.subplots(figsize=(9, 4.5))
    ax.bar(temp[x_col].astype(str), temp[y_col])
    ax.set_title(title)
    ax.tick_params(axis="x", rotation=45)
    plt.tight_layout()
    return fig

# =========================================================
# APP UI
# =========================================================
st.title("ShelfIQ 911")
st.caption("Upload retail data, validate it, auto-map regions from state, run analysis, and download results")

upload_mode = st.radio(
    "Choose upload method",
    ["One Excel workbook", "Separate files"],
    horizontal=True
)

products = None
stores = None
sales_history = None
shelf = None

if upload_mode == "One Excel workbook":
    workbook = st.file_uploader(
        "Upload one Excel workbook with tabs: Sales_History, Products, Stores, and optional Shelf_Snapshot",
        type=["xlsx", "xls"]
    )

    if workbook is not None:
        try:
            products = read_excel_sheet(workbook, "Products")
            stores = read_excel_sheet(workbook, "Stores")
            sales_history = read_excel_sheet(workbook, "Sales_History")

            try:
                shelf = read_excel_sheet(workbook, "Shelf_Snapshot")
            except Exception:
                shelf = None

            st.success("Workbook loaded successfully.")

            with st.expander("Preview detected columns"):
                st.write("Products:", list(products.columns))
                st.write("Stores:", list(stores.columns))
                st.write("Sales_History:", list(sales_history.columns))
                if shelf is not None:
                    st.write("Shelf_Snapshot:", list(shelf.columns))

        except Exception as e:
            st.error(f"Could not read workbook: {e}")

else:
    col1, col2 = st.columns(2)

    with col1:
        sales_file = st.file_uploader("Upload sales history file", type=["csv", "xlsx", "xls"])
        products_file = st.file_uploader("Upload products file", type=["csv", "xlsx", "xls"])

    with col2:
        stores_file = st.file_uploader("Upload stores file", type=["csv", "xlsx", "xls"])
        shelf_file = st.file_uploader("Upload shelf file (optional)", type=["csv", "xlsx", "xls"])

    if sales_file is not None:
        sales_history = read_uploaded_table(sales_file)

    if products_file is not None:
        products = read_uploaded_table(products_file)

    if stores_file is not None:
        stores = read_uploaded_table(stores_file)

    if shelf_file is not None:
        shelf = read_uploaded_table(shelf_file)

run_clicked = st.button("Run ShelfIQ 911 Analysis", type="primary")

if run_clicked:
    if products is None or stores is None or sales_history is None:
        st.error("Please provide Products, Stores, and Sales_History data.")
        st.stop()

    try:
        missing = validate_inputs(
            normalize_columns(products),
            normalize_columns(stores),
            normalize_columns(sales_history),
            normalize_columns(shelf) if shelf is not None else None
        )
        if missing:
            st.error(f"Missing required columns: {missing}")
            st.stop()

        results = run_analysis(products, stores, sales_history, shelf)

        health = results["health_summary"]
        quality = results["quality_checks"]
        recommendations = results["recommendations"]
        sell_in = results["sell_in_opportunities"]
        underperf = results["underperforming_stores"]
        sku = results["sku_velocity_score"]
        dist = results["distribution_gap_index"]
        opp = results["revenue_opportunity_score"]
        declines = results["recent_sku_declines"]
        shelf_df = results["shelf_productivity_score"]
        yoy = results["yoy_growth"]
        momentum = results["momentum"]

        summary = health.iloc[0]

        st.success("Analysis complete.")

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Retail Health", f"{summary['retail_health_score']} ({summary['retail_health_label']})")
        c2.metric("Data Quality", f"{summary['data_quality_score']} ({summary['data_quality_label']})")
        c3.metric("Underperforming Stores", int(summary["underperforming_store_count"]))
        c4.metric("Avg SPI", round(summary["avg_store_performance_index"], 1))
        c5.metric("Revenue Opportunity", f"${summary['estimated_revenue_opportunity']:,.0f}")

        st.divider()

        fail_count = int((quality["status"] == "Fail").sum()) if len(quality) else 0
        warn_count = int((quality["status"] == "Warn").sum()) if len(quality) else 0

        st.subheader("AI Summary")
        st.info(
            f"Retail Health Score is {summary['retail_health_score']} ({summary['retail_health_label']}). "
            f"Data Quality Score is {summary['data_quality_score']} ({summary['data_quality_label']}). "
            f"There are {int(summary['underperforming_store_count'])} underperforming stores and an estimated "
            f"revenue opportunity of ${summary['estimated_revenue_opportunity']:,.0f}. "
            f"Rows uploaded: {int(summary['rows_uploaded'])}, accepted: {int(summary['rows_accepted'])}, rejected: {int(summary['rows_rejected'])}. "
            f"Return Impact Score is {summary['return_impact_score']}%. "
            f"Data quality checks found {fail_count} failures and {warn_count} warnings. "
            f"Regions were auto-assigned from state values in the Stores sheet."
        )

        top_underperf = underperf.sort_values("revenue_opportunity_score", ascending=False).head(5) if len(underperf) else pd.DataFrame()
        top_dist = dist.sort_values("distribution_gap_count", ascending=False).head(5) if len(dist) else pd.DataFrame()
        top_yoy = yoy.sort_values("yoy_sales_growth_pct", ascending=False, na_position="last").head(5) if len(yoy) else pd.DataFrame()
        bottom_yoy = yoy.sort_values("yoy_sales_growth_pct", ascending=True, na_position="last").head(5) if len(yoy) else pd.DataFrame()
        top_momentum = momentum.sort_values("momentum_ratio", ascending=False, na_position="last").head(5) if len(momentum) else pd.DataFrame()
        quality_issues = quality[quality["status"].isin(["Fail", "Warn"])].sort_values(["status", "count"], ascending=[True, False]).head(10) if len(quality) else pd.DataFrame()

        tabs = st.tabs([
            "Executive Summary",
            "Data Quality",
            "Underperforming Stores",
            "SKU Velocity",
            "Distribution Gaps",
            "Recent SKU Declines",
            "Shelf Productivity",
            "YoY Growth",
            "Momentum",
            "Recommendations",
            "Sell-In Opportunities"
        ])

        with tabs[0]:
            st.subheader("Executive Summary")

            e1, e2, e3, e4, e5 = st.columns(5)
            e1.metric("Retail Health", f"{summary['retail_health_score']} ({summary['retail_health_label']})")
            e2.metric("Data Quality", f"{summary['data_quality_score']} ({summary['data_quality_label']})")
            e3.metric("Rows Accepted", int(summary["rows_accepted"]))
            e4.metric("Return Impact", f"{summary['return_impact_score']}%")
            e5.metric("Revenue Opportunity", f"${summary['estimated_revenue_opportunity']:,.0f}")

            st.markdown("### Key Charts")

            chart_col1, chart_col2 = st.columns(2)

            with chart_col1:
                fig = make_bar_chart(
                    underperf.assign(store_label=underperf["store_id"].astype(str)) if len(underperf) else underperf,
                    "store_label",
                    "revenue_opportunity_score",
                    "Top Store Revenue Opportunities",
                    top_n=10,
                    ascending=False
                )
                if fig:
                    st.pyplot(fig)
                else:
                    st.info("No store opportunity chart available.")

            with chart_col2:
                fig = make_bar_chart(
                    dist.assign(gap_label=dist["brand"].astype(str) + " | " + dist["retailer"].astype(str)) if len(dist) else dist,
                    "gap_label",
                    "distribution_gap_count",
                    "Top Distribution Gaps",
                    top_n=10,
                    ascending=False
                )
                if fig:
                    st.pyplot(fig)
                else:
                    st.info("No distribution gap chart available.")

            chart_col3, chart_col4 = st.columns(2)

            with chart_col3:
                if len(yoy):
                    yoy_clean = yoy.dropna(subset=["yoy_sales_growth_pct"]).copy()
                    yoy_clean["sku_label"] = yoy_clean["sku_id"].astype(str)
                    fig = make_bar_chart(
                        yoy_clean,
                        "sku_label",
                        "yoy_sales_growth_pct",
                        "Top YoY Sales Growth",
                        top_n=10,
                        ascending=False
                    )
                    if fig:
                        st.pyplot(fig)
                    else:
                        st.info("No YoY chart available.")
                else:
                    st.info("No YoY chart available.")

            with chart_col4:
                if len(momentum):
                    temp = momentum.copy()
                    temp["sku_label"] = temp["sku_id"].astype(str)
                    fig = make_bar_chart(
                        temp,
                        "sku_label",
                        "momentum_ratio",
                        "Top Momentum Ratios",
                        top_n=10,
                        ascending=False
                    )
                    if fig:
                        st.pyplot(fig)
                    else:
                        st.info("No momentum chart available.")
                else:
                    st.info("No momentum chart available.")

            st.markdown("### Top Underperforming Stores")
            if len(top_underperf):
                cols = [c for c in [
                    "store_id", "retailer", "region", "actual_sales", "expected_sales",
                    "sales_gap", "revenue_opportunity_score", "opportunity_priority", "opportunity_confidence", "exception_flags"
                ] if c in top_underperf.columns]
                st.dataframe(top_underperf[cols], use_container_width=True)
            else:
                st.info("No underperforming stores found.")

            st.markdown("### Top Distribution Gaps")
            if len(top_dist):
                cols = [c for c in ["brand", "category", "retailer", "distribution_gap_count", "distribution_gap_index"] if c in top_dist.columns]
                st.dataframe(top_dist[cols], use_container_width=True)
            else:
                st.info("No distribution gaps found.")

            left, right = st.columns(2)

            with left:
                st.markdown("### Top YoY Winners")
                if len(top_yoy):
                    cols = [c for c in ["sku_id", "brand", "category", "yoy_sales_growth_pct", "yoy_units_growth_pct", "exception_flags"] if c in top_yoy.columns]
                    st.dataframe(top_yoy[cols], use_container_width=True)
                else:
                    st.info("Not enough history for YoY analysis.")

            with right:
                st.markdown("### Top YoY Decliners")
                if len(bottom_yoy):
                    cols = [c for c in ["sku_id", "brand", "category", "yoy_sales_growth_pct", "yoy_units_growth_pct", "exception_flags"] if c in bottom_yoy.columns]
                    st.dataframe(bottom_yoy[cols], use_container_width=True)
                else:
                    st.info("Not enough history for YoY analysis.")

            st.markdown("### Top Momentum Movers")
            if len(top_momentum):
                cols = [c for c in ["sku_id", "brand", "category", "velocity_13w", "velocity_52w", "momentum_ratio", "momentum_flag"] if c in top_momentum.columns]
                st.dataframe(top_momentum[cols], use_container_width=True)
            else:
                st.info("Momentum could not be calculated.")

            st.markdown("### Highest Priority Data Quality Issues")
            if len(quality_issues):
                st.dataframe(quality_issues[["check", "status", "count"]], use_container_width=True)
            else:
                st.success("No data quality failures or warnings found.")

        with tabs[1]:
            st.subheader("Automatic Data Quality Check")
            qc_cols = [c for c in ["check", "status", "count"] if c in quality.columns]
            st.dataframe(quality[qc_cols], use_container_width=True)

        with tabs[2]:
            cols = [c for c in [
                "store_id", "retailer", "region", "actual_sales", "expected_sales",
                "store_performance_index", "sales_gap", "revenue_opportunity_score",
                "opportunity_priority", "opportunity_confidence", "exception_flags"
            ] if c in underperf.columns]
            st.dataframe(underperf[cols].head(100), use_container_width=True)

        with tabs[3]:
            cols = [c for c in [
                "sku_id", "brand", "category", "total_units", "active_stores",
                "velocity_units_per_store_per_week", "sku_velocity_index"
            ] if c in sku.columns]
            st.dataframe(
                sku.sort_values("velocity_units_per_store_per_week", ascending=False)[cols].head(100),
                use_container_width=True
            )

        with tabs[4]:
            cols = [c for c in [
                "brand", "category", "retailer", "current_store_count",
                "retailer_store_universe", "distribution_gap_count", "distribution_gap_index"
            ] if c in dist.columns]
            st.dataframe(
                dist.sort_values("distribution_gap_count", ascending=False)[cols].head(100),
                use_container_width=True
            )

        with tabs[5]:
            cols = [c for c in [
                "sku_id", "brand", "category", "week_end_date", "weekly_sales",
                "prev_week_sales", "wow_change_pct"
            ] if c in declines.columns]
            st.dataframe(declines[cols].head(100), use_container_width=True)

        with tabs[6]:
            if shelf_df is None or len(shelf_df) == 0:
                st.warning("No shelf file was uploaded, so shelf productivity and SEI were not calculated.")
            else:
                cols = [c for c in [
                    "store_id", "retailer", "sku_id", "brand", "category", "facings",
                    "total_sales", "total_units", "shelf_productivity_score",
                    "sales_per_facing", "category_avg_sales_per_facing",
                    "space_efficiency_index", "shelf_action", "exception_flags"
                ] if c in shelf_df.columns]
                st.dataframe(
                    shelf_df.sort_values("space_efficiency_index", ascending=False)[cols].head(100),
                    use_container_width=True
                )

        with tabs[7]:
            if yoy is None or len(yoy) == 0:
                st.warning("Not enough yearly history found to calculate YoY growth. Upload at least 2 years of Sales_History.")
            else:
                cols = [c for c in yoy.columns if c in [
                    "sku_id", "brand", "category", "yoy_sales_growth_pct",
                    "yoy_units_growth_pct", "exception_flags"
                ] or c.startswith("sales_") or c.startswith("units_")]
                st.dataframe(
                    yoy.sort_values("yoy_sales_growth_pct", ascending=False, na_position="last")[cols].head(100),
                    use_container_width=True
                )

        with tabs[8]:
            if momentum is None or len(momentum) == 0:
                st.warning("Momentum could not be calculated.")
            else:
                cols = [c for c in [
                    "sku_id", "brand", "category", "velocity_13w",
                    "velocity_52w", "momentum_ratio", "momentum_flag"
                ] if c in momentum.columns]
                st.dataframe(
                    momentum.sort_values("momentum_ratio", ascending=False, na_position="last")[cols].head(100),
                    use_container_width=True
                )

        with tabs[9]:
            st.subheader("Recommended Actions")
            if len(recommendations):
                st.dataframe(recommendations, use_container_width=True)
            else:
                st.info("No recommendations available.")

        with tabs[10]:
            st.subheader("Sell-In Opportunities")
            if len(sell_in):
                st.dataframe(sell_in, use_container_width=True)
            else:
                st.info("No sell-in opportunities available.")

        st.divider()

        st.download_button(
            "Download Full Results Workbook",
            to_excel_download(results),
            file_name="shelfiq_911_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            "Download Health Summary CSV",
            health.to_csv(index=False).encode("utf-8"),
            file_name="health_summary.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"Analysis failed: {e}")