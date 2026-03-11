
      from io import BytesIO
import pandas as pd
import numpy as np
import streamlit as st

st.set_page_config(page_title="ShelfIQ 911", layout="wide")

# =========================================================
# Helpers
# =========================================================
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]
    return df

def read_excel_sheet(uploaded_file, sheet_name: str) -> pd.DataFrame:
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=0)
    return normalize_columns(df)

def read_uploaded_table(uploaded_file):
    if uploaded_file is None:
        return None

    name = uploaded_file.name.lower()

    if name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
        return normalize_columns(df)

    if name.endswith(".xlsx") or name.endswith(".xls"):
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, header=0)
        return normalize_columns(df)

    raise ValueError(f"Unsupported file type: {uploaded_file.name}")

def validate_inputs(products, stores, sales, shelf=None):
    required_product_cols = {"sku_id"}
    required_store_cols = {"store_id", "format", "retailer", "region", "state"}
    required_sales_cols = {"store_id", "sku_id", "week_end_date", "units"}

    missing = {}

    if not required_product_cols.issubset(products.columns):
        missing["Products"] = sorted(list(required_product_cols - set(products.columns)))

    if not required_store_cols.issubset(stores.columns):
        missing["Stores"] = sorted(list(required_store_cols - set(stores.columns)))

    if not required_sales_cols.issubset(sales.columns):
        missing["Sales_13W"] = sorted(list(required_sales_cols - set(sales.columns)))

    if shelf is not None and len(shelf) > 0:
        required_shelf_cols = {"store_id", "sku_id", "facings", "shelf_share"}
        if not required_shelf_cols.issubset(shelf.columns):
            missing["Shelf_Snapshot"] = sorted(list(required_shelf_cols - set(shelf.columns)))

    return missing

def run_analysis(products, stores, sales, shelf=None):
    products = normalize_columns(products)
    stores = normalize_columns(stores)
    sales = normalize_columns(sales)

    if shelf is None:
        shelf = pd.DataFrame(columns=["store_id", "sku_id", "facings", "shelf_share"])
    else:
        shelf = normalize_columns(shelf)

    missing = validate_inputs(products, stores, sales, shelf if len(shelf) > 0 else None)
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    # Clean types
    sales["week_end_date"] = pd.to_datetime(sales["week_end_date"], errors="coerce")
    sales["units"] = pd.to_numeric(sales["units"], errors="coerce").fillna(0)

    if "sales_dollars" in sales.columns:
        sales["sales_dollars"] = pd.to_numeric(sales["sales_dollars"], errors="coerce").fillna(0)
    elif "sales" in sales.columns:
        sales["sales_dollars"] = pd.to_numeric(sales["sales"], errors="coerce").fillna(0)
    else:
        sales["sales_dollars"] = 0

    if len(shelf) > 0:
        if "facings" in shelf.columns:
            shelf["facings"] = pd.to_numeric(shelf["facings"], errors="coerce").fillna(0)
        if "shelf_share" in shelf.columns:
            shelf["shelf_share"] = pd.to_numeric(shelf["shelf_share"], errors="coerce").fillna(0)

    # Merge
    sales_enriched = (
        sales
        .merge(products, on="sku_id", how="left")
        .merge(stores, on="store_id", how="left")
    )

    for col in ["brand", "category"]:
        if col not in sales_enriched.columns:
            sales_enriched[col] = "Unknown"
        else:
            sales_enriched[col] = sales_enriched[col].fillna("Unknown")

    for col in ["format", "retailer", "region", "state"]:
        if col not in sales_enriched.columns:
            sales_enriched[col] = "Unknown"
        else:
            sales_enriched[col] = sales_enriched[col].fillna("Unknown")

    weeks = max(sales_enriched["week_end_date"].nunique(), 1)

    # =========================================================
    # 1. SKU Velocity
    # =========================================================
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

    # =========================================================
    # 2. Store Performance Index
    # =========================================================
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

    store_perf = store_totals.merge(peer_avg, on=["retailer", "format", "region"], how="left")
    store_perf["expected_sales"] = store_perf["expected_sales"].fillna(store_perf["actual_sales"].mean())
    store_perf["store_performance_index"] = (
        store_perf["actual_sales"] /
        store_perf["expected_sales"].replace(0, np.nan)
    ) * 100
    store_perf["store_performance_index"] = store_perf["store_performance_index"].fillna(0)
    store_perf["sales_gap"] = store_perf["expected_sales"] - store_perf["actual_sales"]
    store_perf["underperforming_flag"] = store_perf["store_performance_index"] < 80

    underperforming_stores = (
        store_perf[store_perf["underperforming_flag"]]
        .sort_values(["sales_gap", "store_performance_index"], ascending=[False, True])
        .reset_index(drop=True)
    )

    # =========================================================
    # 3. Distribution Gap
    # =========================================================
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
        brand_distribution["retailer_store_universe"] - brand_distribution["current_store_count"]
    ).clip(lower=0)

    brand_distribution["distribution_gap_index"] = (
        brand_distribution["distribution_gap_count"] /
        brand_distribution["retailer_store_universe"].replace(0, np.nan)
    ) * 100
    brand_distribution["distribution_gap_index"] = brand_distribution["distribution_gap_index"].fillna(0)

    # =========================================================
    # 4. Revenue Opportunity
    # =========================================================
    store_opportunity = store_perf[[
        "store_id", "retailer", "region", "state", "format",
        "actual_sales", "expected_sales", "sales_gap", "store_performance_index"
    ]].copy()

    store_opportunity["revenue_opportunity_score"] = np.where(
        store_opportunity["sales_gap"] > 0,
        store_opportunity["sales_gap"],
        0
    )

    # =========================================================
    # 5. Recent SKU Declines
    # =========================================================
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

    # =========================================================
    # 6. Shelf Productivity
    # =========================================================
    if len(shelf) > 0 and {"store_id", "sku_id", "facings", "shelf_share"}.issubset(shelf.columns):
        shelf_metrics = (
            shelf
            .merge(products, on="sku_id", how="left")
            .merge(
                sales_enriched.groupby(["store_id", "sku_id"], dropna=False)
                .agg(total_sales=("sales_dollars", "sum"), total_units=("units", "sum"))
                .reset_index(),
                on=["store_id", "sku_id"],
                how="left"
            )
        )
        shelf_metrics["total_sales"] = shelf_metrics["total_sales"].fillna(0)
        shelf_metrics["total_units"] = shelf_metrics["total_units"].fillna(0)
        shelf_metrics["shelf_productivity_score"] = (
            shelf_metrics["total_sales"] / shelf_metrics["facings"].replace(0, np.nan)
        )
        shelf_metrics["shelf_productivity_score"] = shelf_metrics["shelf_productivity_score"].fillna(0)
    else:
        shelf_metrics = pd.DataFrame()

    # =========================================================
    # 7. Health Summary
    # =========================================================
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

    return {
        "health_summary": health_summary,
        "store_performance_index": store_perf,
        "underperforming_stores": underperforming_stores,
        "sku_velocity_score": sku_velocity,
        "distribution_gap_index": brand_distribution,
        "revenue_opportunity_score": store_opportunity,
        "recent_sku_declines": recent_declines,
        "shelf_productivity_score": shelf_metrics,
    }

def to_excel_download(results_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in results_dict.items():
            if df is None or len(df) == 0:
                continue
            df.to_excel(writer, sheet_name=name[:31], index=False)
    output.seek(0)
    return output

# =========================================================
# App
# =========================================================
st.title("ShelfIQ 911")
st.caption("Upload retail data, run analysis, and download results")

upload_mode = st.radio(
    "Choose upload method",
    ["One Excel workbook", "Separate files"],
    horizontal=True
)

products = None
stores = None
sales = None
shelf = None

if upload_mode == "One Excel workbook":
    workbook = st.file_uploader(
        "Upload one Excel workbook with tabs: Sales_13W, Products, Stores, and optional Shelf_Snapshot",
        type=["xlsx", "xls"]
    )

    if workbook is not None:
        try:
            products = read_excel_sheet(workbook, "Products")
            stores = read_excel_sheet(workbook, "Stores")
            sales = read_excel_sheet(workbook, "Sales_13W")

            try:
                shelf = read_excel_sheet(workbook, "Shelf_Snapshot")
            except Exception:
                shelf = None

            st.success("Workbook loaded successfully.")

            with st.expander("Preview detected columns"):
                st.write("Products:", list(products.columns))
                st.write("Stores:", list(stores.columns))
                st.write("Sales_13W:", list(sales.columns))
                if shelf is not None:
                    st.write("Shelf_Snapshot:", list(shelf.columns))

        except Exception as e:
            st.error(f"Could not read workbook: {e}")

else:
    col1, col2 = st.columns(2)

    with col1:
        sales_file = st.file_uploader("Upload sales file", type=["csv", "xlsx", "xls"])
        products_file = st.file_uploader("Upload products file", type=["csv", "xlsx", "xls"])

    with col2:
        stores_file = st.file_uploader("Upload stores file", type=["csv", "xlsx", "xls"])
        shelf_file = st.file_uploader("Upload shelf file (optional)", type=["csv", "xlsx", "xls"])

    if sales_file is not None:
        sales = read_uploaded_table(sales_file)

    if products_file is not None:
        products = read_uploaded_table(products_file)

    if stores_file is not None:
        stores = read_uploaded_table(stores_file)

    if shelf_file is not None:
        shelf = read_uploaded_table(shelf_file)

run_clicked = st.button("Run ShelfIQ 911 Analysis", type="primary")

if run_clicked:
    if products is None or stores is None or sales is None:
        st.error("Please provide Products, Stores, and Sales_13W data.")
        st.stop()

    try:
        results = run_analysis(products, stores, sales, shelf)

        health = results["health_summary"]
        underperf = results["underperforming_stores"]
        sku = results["sku_velocity_score"]
        dist = results["distribution_gap_index"]
        declines = results["recent_sku_declines"]
        shelf_df = results["shelf_productivity_score"]

        summary = health.iloc[0]

        st.success("Analysis complete.")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Retail Health Score", summary["retail_health_score"])
        c2.metric("Underperforming Stores", int(summary["underperforming_store_count"]))
        c3.metric("Average SPI", round(summary["avg_store_performance_index"], 1))
        c4.metric("Revenue Opportunity", f"${summary['estimated_revenue_opportunity']:,.0f}")

        st.divider()

        st.subheader("AI Summary")
        st.info(
            f"Retail Health Score is {summary['retail_health_score']}. "
            f"There are {int(summary['underperforming_store_count'])} underperforming stores and an estimated "
            f"revenue opportunity of ${summary['estimated_revenue_opportunity']:,.0f}. "
            f"Focus first on stores with the biggest sales gaps and SKUs with weak velocity trends."
        )

        tabs = st.tabs([
            "Underperforming Stores",
            "SKU Velocity",
            "Distribution Gaps",
            "Recent SKU Declines",
            "Shelf Productivity"
        ])

        with tabs[0]:
            cols = [c for c in [
                "store_id", "retailer", "region", "actual_sales",
                "expected_sales", "store_performance_index", "sales_gap"
            ] if c in underperf.columns]
            st.dataframe(underperf[cols].head(100), use_container_width=True)
            st.download_button(
                "Download Underperforming Stores CSV",
                underperf.to_csv(index=False).encode("utf-8"),
                file_name="underperforming_stores.csv",
                mime="text/csv"
            )

        with tabs[1]:
            cols = [c for c in [
                "sku_id", "brand", "category", "total_units",
                "active_stores", "velocity_units_per_store_per_week", "sku_velocity_index"
            ] if c in sku.columns]
            st.dataframe(
                sku.sort_values("velocity_units_per_store_per_week", ascending=False)[cols].head(100),
                use_container_width=True
            )
            st.download_button(
                "Download SKU Velocity CSV",
                sku.to_csv(index=False).encode("utf-8"),
                file_name="sku_velocity_score.csv",
                mime="text/csv"
            )

        with tabs[2]:
            cols = [c for c in [
                "brand", "category", "retailer", "current_store_count",
                "retailer_store_universe", "distribution_gap_count", "distribution_gap_index"
            ] if c in dist.columns]
            st.dataframe(
                dist.sort_values("distribution_gap_count", ascending=False)[cols].head(100),
                use_container_width=True
            )
            st.download_button(
                "Download Distribution Gaps CSV",
                dist.to_csv(index=False).encode("utf-8"),
                file_name="distribution_gap_index.csv",
                mime="text/csv"
            )

        with tabs[3]:
            cols = [c for c in [
                "sku_id", "brand", "category", "week_end_date",
                "weekly_sales", "prev_week_sales", "wow_change_pct"
            ] if c in declines.columns]
            st.dataframe(declines[cols].head(100), use_container_width=True)
            st.download_button(
                "Download Recent SKU Declines CSV",
                declines.to_csv(index=False).encode("utf-8"),
                file_name="recent_sku_declines.csv",
                mime="text/csv"
            )

        with tabs[4]:
            if shelf_df is None or len(shelf_df) == 0:
                st.warning("No shelf file was uploaded, so shelf productivity was not calculated.")
            else:
                cols = [c for c in [
                    "store_id", "sku_id", "brand", "category",
                    "facings", "total_sales", "total_units", "shelf_productivity_score"
                ] if c in shelf_df.columns]
                st.dataframe(
                    shelf_df.sort_values("shelf_productivity_score", ascending=False)[cols].head(100),
                    use_container_width=True
                )
                st.download_button(
                    "Download Shelf Productivity CSV",
                    shelf_df.to_csv(index=False).encode("utf-8"),
                    file_name="shelf_productivity_score.csv",
                    mime="text/csv"
                )

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