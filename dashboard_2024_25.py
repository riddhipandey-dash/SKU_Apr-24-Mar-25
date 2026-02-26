import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from pathlib import Path

# --- Config & Style ---
st.set_page_config(page_title="SKU Sales Dashboard (Apr,24 - Mar,25)", layout="wide")
st.markdown("""
    <style>
    .stApp {
        background: linear-gradient(120deg, #232526 0%, #414345 100%) fixed;
        min-height: 100vh;
    }
    .block-container {
        background: #232526;
        border-radius: 18px;
        box-shadow: 0 8px 32px 0 rgba(20, 20, 20, 0.25);
        padding: 2.5rem 2.5rem 2.5rem 2.5rem;
        margin-top: 2rem;
    }
    h1, h2, h3, h4, h5, h6 {
        color: #fff !important;
        font-weight: 800;
        margin-bottom: 1.2rem;
        letter-spacing: 0.5px;
    }
    /* Custom metric card styling */
    div[data-testid="stMetric"] {
        background: #11131a;
        border-radius: 18px;
        padding: 18px 0 14px 0;
        box-shadow: 0 4px 24px 0 rgba(20, 20, 20, 0.25);
        border: 1.5px solid #333;
        margin-bottom: 0.5rem;
        color: #fff !important;
    }
    div[data-testid="stMetric"] > label, div[data-testid="stMetric"] > div {
        color: #fff !important;
    }
    .stSidebar {
        background: #232526;
        color: #fff;
        border-radius: 0 18px 18px 0;
    }
    /* Removed custom filter box background and accent colors */
    .sidebar-filter-box {
        border-radius: 14px;
        margin-bottom: 1.2rem;
        padding: 1.1rem 1rem 0.7rem 1rem;
        box-shadow: 0 2px 12px 0 rgba(0,0,0,0.10);
        background: #232526;
        border-left: none;
    }
    .sidebar-filter-month,
    .sidebar-filter-state,
    .sidebar-filter-category,
    .sidebar-filter-sku {
        background: none;
        border-left: none;
    }
    </style>
""", unsafe_allow_html=True)

# --- Data Loading ---
BASE_DIR = Path(__file__).parent
DATA_PATH = BASE_DIR / "REGION & SKU WISE SALE(2024-2025).xlsx"
if not DATA_PATH.exists():
    st.error("Excel file not found in repository.")
    st.stop()

@st.cache_data
def load_and_clean_data():
    xls = pd.ExcelFile(DATA_PATH)
    all_detail, totals_rows, month_labels = [], [], {}
    for sheet in xls.sheet_names:
        raw = pd.read_excel(DATA_PATH, sheet_name=sheet, header=None)
        month_label = raw.iloc[1, 0] if isinstance(raw.iloc[1, 0], str) else sheet
        month_labels[sheet] = month_label
        header = [str(c).strip() for c in raw.iloc[2].tolist()]
        data = raw.drop(index=3).iloc[4:].copy()
        data.columns = header
        data = data.dropna(axis=0, how="all").reset_index(drop=True)
        data["S.NO."] = data["S.NO."].astype(str).str.strip()
        data["PARTY NAME"] = data["PARTY NAME"].astype(str).str.strip()
        is_total = (
            data["S.NO."].str.upper().eq("TOTAL") |
            data["PARTY NAME"].str.upper().eq("TOTAL")
        )
        totals_row = data[is_total].copy()
        detail = data.copy()
        meta_cols = ["S.NO.", "PARTY NAME", "TOWN", "STATE"]
        for col in meta_cols:
            if col not in detail.columns:
                detail[col] = ""
        numeric_cols = [c for c in detail.columns if c not in meta_cols]
        for col in numeric_cols:
            detail[col] = pd.to_numeric(detail[col], errors="coerce").fillna(0)
            totals_row[col] = pd.to_numeric(totals_row[col], errors="coerce").fillna(0)
        detail["Month"] = sheet
        totals_row["Month"] = sheet
        all_detail.append(detail)
        totals_rows.append(totals_row)
    detail = pd.concat(all_detail, ignore_index=True) if all_detail else pd.DataFrame()
    totals_row = pd.concat(totals_rows, ignore_index=True) if totals_rows else pd.DataFrame()
    meta_cols = ["S.NO.", "PARTY NAME", "TOWN", "STATE", "Month"]
    category_cols = [
        c for c in detail.columns if c not in meta_cols
        if str(c).strip().upper().endswith("TOTAL") and str(c).strip().upper() not in {"TOTAL", "TOTAL AMOUNT"}
    ]
    return detail, totals_row, month_labels, category_cols, xls.sheet_names

detail, totals_row, month_labels, category_cols, month_order = load_and_clean_data()

META_COLS = ["S.NO.", "PARTY NAME", "TOWN", "STATE", "Month"]
# --- Sidebar Filters ---
def sidebar_filters():
    st.sidebar.header("Filters")
    with st.sidebar:
        st.markdown("<div class='sidebar-filter-box sidebar-filter-month'>", unsafe_allow_html=True)
        selected_months = st.multiselect("Select Month(s)", month_order, default=month_order)
        st.markdown("</div>", unsafe_allow_html=True)
        state_options = sorted(detail["STATE"].dropna().unique().tolist()) if "STATE" in detail.columns else []
        st.markdown("<div class='sidebar-filter-box sidebar-filter-state'>", unsafe_allow_html=True)
        selected_states = st.multiselect("Select State(s)", state_options, default=state_options)
        st.markdown("</div>", unsafe_allow_html=True)
        category_options = sorted(category_cols) if category_cols else []
        st.markdown("<div class='sidebar-filter-box sidebar-filter-category'>", unsafe_allow_html=True)
        selected_categories = st.multiselect("Select Category(ies)", category_options, default=category_options)
        st.markdown("</div>", unsafe_allow_html=True)
        sku_options = sorted([c for c in detail.columns if c not in META_COLS and c not in category_cols and c != "TOTAL AMOUNT"])
        st.markdown("<div class='sidebar-filter-box sidebar-filter-sku'>", unsafe_allow_html=True)
        selected_skus = st.multiselect("Select SKU(s)", sku_options, default=sku_options)
        st.markdown("</div>", unsafe_allow_html=True)
    filtered_detail = detail.copy()
    if selected_months:
        filtered_detail = filtered_detail[filtered_detail["Month"].isin(selected_months)]
    if selected_states and "STATE" in filtered_detail.columns:
        filtered_detail = filtered_detail[filtered_detail["STATE"].isin(selected_states)]
    # Filter by category columns (if any selected)
    if selected_categories and category_cols:
        filtered_detail = filtered_detail.copy()
        filtered_detail = filtered_detail.loc[:, META_COLS + selected_categories + [c for c in filtered_detail.columns if c not in META_COLS + category_cols]]
    # Filter by SKU columns (if any selected)
    if selected_skus:
        filtered_detail = filtered_detail.copy()
        filtered_detail = filtered_detail.loc[:, META_COLS + selected_skus + [c for c in filtered_detail.columns if c not in META_COLS + sku_options]]
    return filtered_detail

filtered_detail = sidebar_filters()

# --- Dashboard Heading ---
st.markdown("<h1 style='text-align:center;'>SKU Sales Dashboard <span style='font-size:1.2rem;'>(Apr,24 - Mar,25)</span></h1>", unsafe_allow_html=True)

# --- Metrics Section ---
META_COLS = ["S.NO.", "PARTY NAME", "TOWN", "STATE", "Month"]
numeric_cols = [c for c in detail.columns if c not in META_COLS]
money_col = "TOTAL AMOUNT" if "TOTAL AMOUNT" in detail.columns else None
sku_cols = [c for c in numeric_cols if c not in category_cols and c != money_col and str(c).strip().upper() not in {"TOTAL", "TOTAL AMOUNT"}]
case_cols = category_cols if category_cols else sku_cols

total_cases_by_month = {}
for month in month_order:
    month_total_row = totals_row[totals_row["Month"] == month]
    if not month_total_row.empty:
        valid_case_cols = [c for c in case_cols if c in month_total_row.columns]
        total_cases_by_month[month] = month_total_row[valid_case_cols].sum(axis=1).sum()
    else:
        total_cases_by_month[month] = 0
if "TOTAL AMOUNT" in totals_row.columns:
    total_amount_by_month = totals_row.set_index("Month")["TOTAL AMOUNT"].reindex(month_order).fillna(0)
    grand_total_amount = float(total_amount_by_month.sum())
else:
    grand_total_amount = 0
sku_cols_for_table = [c for c in sku_cols if c in totals_row.columns]
sku_sum_by_month = totals_row.set_index("Month")[sku_cols_for_table].sum(axis=1).reindex(month_order).fillna(0)
grand_total_cases = int(sku_sum_by_month.sum())
parties = filtered_detail["PARTY NAME"].nunique() if "PARTY NAME" in filtered_detail.columns else 0

# Use markdown for clear, large metric values with dark background
col1, col2, col3 = st.columns(3)
with col1:
    st.markdown(f"""
        <div style='background:#11131a;padding:2rem 0;border-radius:18px;text-align:center;box-shadow:0 4px 24px 0 rgba(20,20,20,0.25);border:1.5px solid #333;'>
            <div style='color:#fff;font-size:1.1rem;font-weight:700;'>Total Amount (₹)</div>
            <div style='color:#00e676;font-size:2.2rem;font-weight:900;margin-top:0.5rem;'>₹{grand_total_amount:,.0f}</div>
        </div>
    """, unsafe_allow_html=True)
with col2:
    st.markdown(f"""
        <div style='background:#11131a;padding:2rem 0;border-radius:18px;text-align:center;box-shadow:0 4px 24px 0 rgba(20,20,20,0.25);border:1.5px solid #333;'>
            <div style='color:#fff;font-size:1.1rem;font-weight:700;'>Total Cases</div>
            <div style='color:#00b0ff;font-size:2.2rem;font-weight:900;margin-top:0.5rem;'>{grand_total_cases:,}</div>
        </div>
    """, unsafe_allow_html=True)
with col3:
    st.markdown(f"""
        <div style='background:#11131a;padding:2rem 0;border-radius:18px;text-align:center;box-shadow:0 4px 24px 0 rgba(20,20,20,0.25);border:1.5px solid #333;'>
            <div style='color:#fff;font-size:1.1rem;font-weight:700;'>Parties</div>
            <div style='color:#ffd600;font-size:2.2rem;font-weight:900;margin-top:0.5rem;'>{parties:,}</div>
        </div>
    """, unsafe_allow_html=True)

# --- Monthly Category Distribution ---
if "TOTAL AMOUNT" in totals_row.columns:
    valid_category_cols = [c for c in category_cols if c in totals_row.columns]
    if valid_category_cols:
        monthly_cat = totals_row.set_index("Month")[valid_category_cols].reindex(month_order).fillna(0)
        monthly_cat.index = pd.CategoricalIndex(monthly_cat.index, categories=month_order, ordered=True)
        monthly_cat = monthly_cat.sort_index()
        melted = monthly_cat.reset_index().melt(id_vars="Month", var_name="Category", value_name="Cases")
        fig_cat_dist = px.bar(
            melted, x="Month", y="Cases", color="Category", barmode="stack", title=None,
            color_discrete_sequence=px.colors.qualitative.Set2
        )
        fig_cat_dist.update_layout(
            plot_bgcolor='#232526', paper_bgcolor='#232526', font_color='#fff',
            xaxis=dict(title_font=dict(color='#fff'), tickfont=dict(color='#fff')),
            yaxis=dict(title_font=dict(color='#fff'), tickfont=dict(color='#fff')),
            legend=dict(font=dict(color='#fff'))
        )
        st.markdown("<h3 style='color:black;'>Monthly Category Distribution (Apr,24 - Mar,25)</h3>", unsafe_allow_html=True)
        st.plotly_chart(fig_cat_dist, use_container_width=True, key="cat_dist")
        st.markdown("This stacked bar chart shows the monthly distribution of cases by product category, helping you identify seasonal trends and category performance over time.")

# --- Monthly Total Amount ---
if money_col:
    monthly = filtered_detail.groupby("Month", as_index=False)[money_col].sum()
    if "TOTAL AMOUNT" in totals_row.columns:
        total_amount_by_month = totals_row.set_index("Month")["TOTAL AMOUNT"].reindex(month_order).fillna(0)
        amount_chart_df = pd.DataFrame({"Month": month_order, "TOTAL AMOUNT": total_amount_by_month.values})
        amount_chart_df["Month"] = pd.Categorical(amount_chart_df["Month"], categories=month_order, ordered=True)
        amount_chart_df = amount_chart_df.sort_values("Month")
        fig_month = px.line(
            amount_chart_df, x="Month", y="TOTAL AMOUNT", markers=True, title=None,
            line_shape='spline',
            color_discrete_sequence=['#00e676']
        )
        fig_month.update_traces(line=dict(color='#00e676', width=4), marker=dict(color='#00e676', size=10))
        fig_month.update_layout(
            plot_bgcolor='#232526', paper_bgcolor='#232526', font_color='#fff',
            xaxis=dict(title_font=dict(color='#fff'), tickfont=dict(color='#fff')),
            yaxis=dict(title_font=dict(color='#fff'), tickfont=dict(color='#fff')),
            legend=dict(font=dict(color='#fff'))
        )
        st.markdown("<h3 style='color:black;'>Monthly Total Amount (₹) (Apr,24 - Mar,25)</h3>", unsafe_allow_html=True)
        st.plotly_chart(fig_month, use_container_width=True, key="month")
        st.markdown("This line chart tracks the total sales amount each month, highlighting peaks and dips in revenue throughout the year.")

# --- Category Analysis ---
category_totals = filtered_detail[category_cols].sum().sort_values(ascending=False) if category_cols else pd.Series(dtype=float)
if not category_totals.empty:
    fig_prod = px.bar(
        category_totals.head(10),
        x=category_totals.head(10).values,
        y=category_totals.head(10).index,
        orientation="h",
        title=None,
        labels={"x": "Cases", "y": "Category"},
        color_discrete_sequence=['#00b0ff']
    )
    fig_prod.update_layout(
        plot_bgcolor='#232526', paper_bgcolor='#232526', font_color='#fff',
        xaxis=dict(title_font=dict(color='#fff'), tickfont=dict(color='#fff')),
        yaxis=dict(title_font=dict(color='#fff'), tickfont=dict(color='#fff')),
        legend=dict(font=dict(color='#fff'))
    )
    st.markdown("<h3 style='color:black;'>Top 10 Categories by Cases (Apr,24 - Mar,25)</h3>", unsafe_allow_html=True)
    st.plotly_chart(fig_prod, use_container_width=True, key="prod_top")
    st.markdown("These are the top 10 product categories by total cases sold, revealing your best-performing segments.")
    fig_prod_bottom = px.bar(
        category_totals.tail(10),
        x=category_totals.tail(10).values,
        y=category_totals.tail(10).index,
        orientation="h",
        title=None,
        labels={"x": "Cases", "y": "Category"},
        color_discrete_sequence=['#ffd600']
    )
    fig_prod_bottom.update_layout(
        plot_bgcolor='#232526', paper_bgcolor='#232526', font_color='#fff',
        xaxis=dict(title_font=dict(color='#fff'), tickfont=dict(color='#fff')),
        yaxis=dict(title_font=dict(color='#fff'), tickfont=dict(color='#fff')),
        legend=dict(font=dict(color='#fff'))
    )
    st.markdown("<h3 style='color:black;'>Bottom 10 Categories by Cases (Apr,24 - Mar,25)</h3>", unsafe_allow_html=True)
    st.plotly_chart(fig_prod_bottom, use_container_width=True, key="prod_bottom")
    st.markdown("These are the bottom 10 product categories by total cases sold, highlighting underperforming segments that may need attention.")
    fig_cat = px.pie(
        values=category_totals.values,
        names=category_totals.index,
        hole=0.5,
        title=None,
        color_discrete_sequence=px.colors.sequential.Purples_r
    )
    fig_cat.update_layout(
        plot_bgcolor='#232526', paper_bgcolor='#232526', font_color='#fff',
        legend=dict(font=dict(color='#fff'))
    )
    st.markdown("<h3 style='color:black;'>Category Share (Totals) (Apr,24 - Mar,25)</h3>", unsafe_allow_html=True)
    st.plotly_chart(fig_cat, use_container_width=True, key="cat_share")
    st.markdown("This donut chart shows the overall share of each category in total cases sold, giving a quick view of your product mix.")

# --- Top SKUs ---
st.header("Top SKUs by Cases (Apr,24 - Mar,25)")
sku_cols_for_table = [c for c in sku_cols if c in filtered_detail.columns]
if sku_cols_for_table:
    sku_totals = filtered_detail[sku_cols_for_table].sum().sort_values(ascending=False)
    sku_df = sku_totals.head(15).reset_index()
    sku_df.columns = ['SKU', 'Cases']
    fig_sku = px.bar(
        sku_df,
        x='Cases',
        y='SKU',
        orientation="h",
        title=None,
        labels={"Cases": "Cases", "SKU": "SKU"},
        color='SKU',
        color_discrete_sequence=px.colors.qualitative.Set3
    )
    fig_sku.update_layout(
        plot_bgcolor='#232526', paper_bgcolor='#232526', font_color='#fff',
        xaxis=dict(title_font=dict(color='#fff'), tickfont=dict(color='#fff')),
        yaxis=dict(title_font=dict(color='#fff'), tickfont=dict(color='#fff')),
        legend=dict(font=dict(color='#fff'))
    )
    st.markdown("<h3 style='color:black;'>Top 15 SKUs by Cases (Apr,24 - Mar,25)</h3>", unsafe_allow_html=True)
    st.plotly_chart(fig_sku, use_container_width=True, key="sku_top")
    st.markdown("This bar chart displays the top 15 SKUs by cases sold, helping you spot your star products.")
else:
    st.info("No SKU columns detected. Add SKU columns or confirm your column naming.")

# --- New Product Introductions ---
st.header("New Product Introductions (Apr,24 - Mar,25)")
sku_cols_for_table = [c for c in sku_cols if c in totals_row.columns]
monthly_sku_presence = totals_row.set_index("Month")[sku_cols_for_table].reindex(month_order).fillna(0)
monthly_sku_presence.index = pd.CategoricalIndex(monthly_sku_presence.index, categories=month_order, ordered=True)
monthly_sku_presence = monthly_sku_presence.sort_index()
sku_introduced = []
seen_skus = set()
for month in monthly_sku_presence.index:
    current_month_skus = set(monthly_sku_presence.columns[monthly_sku_presence.loc[month] > 0])
    new_this_month = current_month_skus - seen_skus
    sku_introduced.append(len(new_this_month))
    seen_skus.update(current_month_skus)
sku_intro_df = pd.DataFrame({"Month": list(monthly_sku_presence.index), "New SKUs Introduced": sku_introduced})
sku_intro_df["Month"] = pd.Categorical(sku_intro_df["Month"], categories=month_order, ordered=True)
sku_intro_df = sku_intro_df.sort_values("Month")
fig_new_sku = px.line(
    sku_intro_df,
    x="Month",
    y="New SKUs Introduced",
    markers=True,
    title=None,
    line_shape='spline',
    color_discrete_sequence=['#ffd600']
)
fig_new_sku.update_traces(line=dict(color='#ffd600', width=4), marker=dict(color='#ffd600', size=10))
fig_new_sku.update_layout(
    plot_bgcolor='#232526', paper_bgcolor='#232526', font_color='#fff',
    xaxis=dict(title_font=dict(color='#fff'), tickfont=dict(color='#fff')),
    yaxis=dict(title_font=dict(color='#fff'), tickfont=dict(color='#fff')),
    legend=dict(font=dict(color='#fff'))
)
st.markdown("<h3 style='color:black;'>New SKUs Introduced Each Month (Apr,24 - Mar,25)</h3>", unsafe_allow_html=True)
st.plotly_chart(fig_new_sku, use_container_width=True, key="sku_new")
st.markdown("This line chart shows how many new SKUs were introduced each month, indicating innovation and portfolio expansion.")
sku_intro_names = []
seen_skus = set()
for month in monthly_sku_presence.index:
    current_month_skus = set(monthly_sku_presence.columns[monthly_sku_presence.loc[month] > 0])
    new_this_month = current_month_skus - seen_skus
    sku_intro_names.append((month, sorted(new_this_month)))
    seen_skus.update(current_month_skus)
st.markdown("**Summary of SKUs Introduced Each Month:**")
for month, skus in sku_intro_names:
    if skus:
        st.markdown(f"- **{month}:** {', '.join(skus)}")
st.markdown("Above is a month-wise list of new SKUs introduced, useful for tracking launches and adoption.")

# --- State-wise Performance ---
st.header("State-wise Performance (Apr,24 - Mar,25)")
if "STATE" in filtered_detail.columns and money_col:
    # Ensure money_col is a string and column exists
    if isinstance(money_col, str) and money_col in filtered_detail.columns:
        state = filtered_detail.groupby("STATE", as_index=False)[money_col].sum()
    #    state = state.sort_values(by="TOTAL AMOUNT", ascending=False).head(12)
    else:
        state = pd.DataFrame(columns=["STATE", str(money_col)])
    fig_state = px.bar(
        state,
        x=money_col,
        y="STATE",
        orientation="h",
        title=None,
        color_discrete_sequence=['#43a047']
    )
    fig_state.update_layout(
        plot_bgcolor='#232526', paper_bgcolor='#232526', font_color='#fff',
        xaxis=dict(title_font=dict(color='#fff'), tickfont=dict(color='#fff')),
        yaxis=dict(title_font=dict(color='#fff'), tickfont=dict(color='#fff')),
        legend=dict(font=dict(color='#fff'))
    )
    st.markdown("<h3 style='color:black;'>Top States by Total Amount (₹) (Apr,24 - Mar,25)</h3>", unsafe_allow_html=True)
    st.plotly_chart(fig_state, use_container_width=True, key="state")
    st.markdown("This horizontal bar chart ranks states by total sales amount, helping you identify your strongest and weakest markets.")
