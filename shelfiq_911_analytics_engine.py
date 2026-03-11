from pathlib import Path
import pandas as pd
import numpy as np

# =====================================
# ShelfIQ 911 Analytics Engine
# =====================================

# Current folder
BASE_PATH = Path(".")

# Excel workbook name
WORKBOOK_PATH = BASE_PATH / "shelfiq_911_dataset_download.xlsx"

# Output folder
OUTPUT_DIR = BASE_PATH / "shelfiq_911_outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

print("Starting ShelfIQ 911 analytics engine...")
print(f"Looking for workbook at: {WORKBOOK_PATH.resolve()}")

if not WORKBOOK_PATH.exists():
    print("\nERROR: Excel file not found.")
    print("Make sure this file is in the same folder as this script:")
    print("shelfiq_911_dataset_download.xlsx")
    raise FileNotFoundError(f"Missing file: {WORKBOOK_PATH}")

# -----------------------------
# Load data
# -----------------------------
products = pd.read_excel(WORKBOOK_PATH, sheet_name="Products")
stores = pd.read_excel(WORKBOOK_PATH, sheet_name="Stores")
sales = pd.read_excel(WORKBOOK_PATH, sheet_name="Sales_13W")
shelf = pd.read_excel(WORKBOOK_PATH, sheet_name="Shelf_Snapshot")

# -----------------------------
# Standardize columns
# -----------------------------
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]
    return df

products = normalize_columns(products)
stores = normalize_columns(stores)
sales = normalize_columns(sales)
shelf = normalize_columns(shelf)

# -----------------------------
# Validate columns
# -----------------------------
required_product_cols = {"sku_id"}
required_store_cols = {"store_id", "retailer", "region", "state", "format"}
required_sales_cols = {"week_end_date", "store_id", "sku_id", "units"}
required_shelf_cols = {"store_id", "sku_id", "facings"}

if not required_product_cols.issubset(products.columns):
    raise ValueError(f"Products sheet is missing columns: {required_product_cols - set(products.columns)}")

if not required_store_cols.issubset(stores.columns):
    raise ValueError(f"Stores sheet is missing columns: {required_store_cols - set(stores.columns)}")

if not required_sales_cols.issubset(sales.columns):
    raise ValueError(f"Sales_13W sheet is missing columns: {required_sales_cols - set(sales.columns)}")

if not required_shelf_cols.issubset(shelf.columns):
    raise ValueError(f"Shelf_Snapshot sheet is missing columns: {required_shelf_cols - set(shelf.columns)}")

# -----------------------------
# Basic cleanup
# -----------------------------
sales["week_end_date"] = pd.to_datetime(sales["week_end_date"], errors="coerce")

numeric_candidates = ["units", "sales_dollars", "sales", "facings", "shelf_share"]
for col in numeric_candidates:
    if col in sales.columns:
        sales[col] = pd.to_numeric(sales[col], errors="coerce")
    if col in shelf.columns:
        shelf[col] = pd.to_numeric(shelf[col], errors="coerce")

# Support either "sales_dollars" or "sales"
if "sales_dollars" not in sales.columns:
    if "sales" in sales.columns:
        sales["sales_dollars"] = sales["sales"]
    else:
        raise ValueError("Sales_13W sheet must contain either 'sales_dollars' or 'sales' column.")

# Fill nulls
sales["units"] = sales["units"].fillna(0)
sales["sales_dollars"] = sales["sales_dollars"].fillna(0)
shelf["facings"] = shelf["facings"].fillna(0)

# -----------------------------
# Join master data
# -----------------------------
sales_enriched = (
    sales
    .merge(products, on="sku_id", how="left")
    .merge(stores, on="store_id", how="left")
)

# Fill missing descriptive values
for col in ["brand", "category"]:
    if col not in sales_enriched.columns:
        sales_enriched[col] = "Unknown"
    else:
        sales_enriched[col] = sales_enriched[col].fillna("Unknown")

for col in ["retailer", "region", "state", "format"]:
    if col not in sales_enriched.columns:
        sales_enriched[col] = "Unknown"
    else:
        sales_enriched[col] = sales_enriched[col].fillna("Unknown")

# -----------------------------
# Core metrics
# -----------------------------
weeks = max(sales_enriched["week_end_date"].nunique(), 1)

# 1) SKU Velocity Score
sku_velocity = (
    sales_enriched
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
    weeks
)

category_avg_velocity = (
    sku_velocity
    .groupby("category", dropna=False)["velocity_units_per_store_per_week"]
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

# 2) Store Performance Index
store_totals = (
    sales_enriched
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

store_perf = store_totals.merge(
    peer_avg,
    on=["retailer", "format", "region"],
    how="left"
)

store_perf["expected_sales"] = store_perf["expected_sales"].fillna(store_perf["actual_sales"].mean())
store_perf["store_performance_index"] = (
    store_perf["actual_sales"] /
    store_perf["expected_sales"].replace(0, np.nan)
) * 100
store_perf["store_performance_index"] = store_perf["store_performance_index"].fillna(0)
store_perf["sales_gap"] = store_perf["expected_sales"] - store_perf["actual_sales"]
store_perf["underperforming_flag"] = store_perf["store_performance_index"] < 80

# 3) Distribution Gap Index
carried = (
    sales_enriched
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
    brand_distribution["retailer_store_universe"] -
    brand_distribution["current_store_count"]
).clip(lower=0)

brand_distribution["distribution_gap_index"] = (
    brand_distribution["distribution_gap_count"] /
    brand_distribution["retailer_store_universe"].replace(0, np.nan)
) * 100
brand_distribution["distribution_gap_index"] = brand_distribution["distribution_gap_index"].fillna(0)

# 4) Shelf Productivity Score
shelf_metrics = (
    shelf
    .merge(products, on="sku_id", how="left")
    .merge(
        sales_enriched
        .groupby(["store_id", "sku_id"], dropna=False)
        .agg(
            total_sales=("sales_dollars", "sum"),
            total_units=("units", "sum")
        )
        .reset_index(),
        on=["store_id", "sku_id"],
        how="left"
    )
)

if "brand" not in shelf_metrics.columns:
    shelf_metrics["brand"] = "Unknown"
else:
    shelf_metrics["brand"] = shelf_metrics["brand"].fillna("Unknown")

if "category" not in shelf_metrics.columns:
    shelf_metrics["category"] = "Unknown"
else:
    shelf_metrics["category"] = shelf_metrics["category"].fillna("Unknown")

shelf_metrics["total_sales"] = shelf_metrics["total_sales"].fillna(0)
shelf_metrics["total_units"] = shelf_metrics["total_units"].fillna(0)

shelf_metrics["shelf_productivity_score"] = (
    shelf_metrics["total_sales"] / shelf_metrics["facings"].replace(0, np.nan)
)
shelf_metrics["shelf_productivity_score"] = shelf_metrics["shelf_productivity_score"].fillna(0)

# 5) Revenue Opportunity Score
store_opportunity = store_perf[[
    "store_id",
    "retailer",
    "region",
    "state",
    "format",
    "actual_sales",
    "expected_sales",
    "sales_gap",
    "store_performance_index"
]].copy()

store_opportunity["revenue_opportunity_score"] = np.where(
    store_opportunity["sales_gap"] > 0,
    store_opportunity["sales_gap"],
    0
)

# -----------------------------
# Alerts
# -----------------------------
underperforming_stores = (
    store_perf[store_perf["underperforming_flag"]]
    .sort_values(["sales_gap", "store_performance_index"], ascending=[False, True])
    .reset_index(drop=True)
)

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

# -----------------------------
# Retail Health Score
# -----------------------------
underperf_rate = float(store_perf["underperforming_flag"].mean()) if len(store_perf) else 0
avg_spi = float(store_perf["store_performance_index"].fillna(100).mean()) if len(store_perf) else 100
dist_gap_rate = float(brand_distribution["distribution_gap_index"].fillna(0).mean()) if len(brand_distribution) else 0

a = min(max(avg_spi, 0), 120) / 120 * 50
b = (1 - min(max(underperf_rate, 0), 1)) * 30
c = (1 - min(max(dist_gap_rate / 100, 0), 1)) * 20

retail_health_score = round(a + b + c, 1)

health_summary = pd.DataFrame([{
    "retail_health_score": retail_health_score,
    "store_count": int(store_perf["store_id"].nunique()),
    "sku_count": int(products["sku_id"].nunique()),
    "underperforming_store_count": int(store_perf["underperforming_flag"].sum()),
    "avg_store_performance_index": round(avg_spi, 2),
    "avg_distribution_gap_index": round(dist_gap_rate, 2),
    "estimated_revenue_opportunity": round(float(store_opportunity["revenue_opportunity_score"].sum()), 2),
}])

# -----------------------------
# Export outputs
# -----------------------------
health_summary.to_csv(OUTPUT_DIR / "health_summary.csv", index=False)
store_perf.to_csv(OUTPUT_DIR / "store_performance_index.csv", index=False)
underperforming_stores.to_csv(OUTPUT_DIR / "underperforming_stores.csv", index=False)
sku_velocity.to_csv(OUTPUT_DIR / "sku_velocity_score.csv", index=False)
brand_distribution.to_csv(OUTPUT_DIR / "distribution_gap_index.csv", index=False)
store_opportunity.to_csv(OUTPUT_DIR / "revenue_opportunity_score.csv", index=False)
recent_declines.to_csv(OUTPUT_DIR / "recent_sku_declines.csv", index=False)
shelf_metrics.to_csv(OUTPUT_DIR / "shelf_productivity_score.csv", index=False)

# -----------------------------
# Console summary
# -----------------------------
print("\nShelfIQ 911 analytics engine complete.")
print(f"Retail Health Score: {retail_health_score}")
print(f"Outputs saved to: {OUTPUT_DIR.resolve()}")

print("\nTop underperforming stores:")
if underperforming_stores.empty:
    print("No underperforming stores found.")
else:
    print(
        underperforming_stores.head(10)[[
            "store_id",
            "retailer",
            "region",
            "actual_sales",
            "expected_sales",
            "store_performance_index",
            "sales_gap"
        ]]
    )

print("\nDone.")