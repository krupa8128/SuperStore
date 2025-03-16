import streamlit as st
import pandas as pd
import plotly.express as px

# Load data
file_path = "Sample - Superstore-1.xlsx"
xls = pd.ExcelFile(file_path)
orders_df = pd.read_excel(xls, sheet_name="Orders")
returns_df = pd.read_excel(xls, sheet_name="Returns")
people_df = pd.read_excel(xls, sheet_name="People")

# Merge return data
orders_df["Returned"] = orders_df["Order ID"].isin(returns_df["Order ID"])

# Convert Order Date to datetime
orders_df["Order Date"] = pd.to_datetime(orders_df["Order Date"])

# Streamlit UI
st.set_page_config(page_title="Superstore Dashboard", layout="wide")
st.title("📊 Superstore Business Performance Dashboard")

# Sidebar Filters
st.sidebar.header("Filter Options")
selected_region = st.sidebar.multiselect("Select Region", options=orders_df["Region"].unique(), default=orders_df["Region"].unique())
selected_category = st.sidebar.multiselect("Select Category", options=orders_df["Category"].unique(), default=orders_df["Category"].unique())

# Filtered Data
df_filtered = orders_df[(orders_df["Region"].isin(selected_region)) & (orders_df["Category"].isin(selected_category))]

# KPIs
col1, col2, col3 = st.columns(3)
with col1:
    total_sales = df_filtered["Sales"].sum()
    st.metric("Total Sales", f"${total_sales:,.2f}")
with col2:
    total_profit = df_filtered["Profit"].sum()
    st.metric("Total Profit", f"${total_profit:,.2f}")
with col3:
    return_rate = df_filtered["Returned"].mean() * 100
    st.metric("Return Rate", f"{return_rate:.2f}%")

# Sales by Region
fig_sales_region = px.bar(df_filtered.groupby("Region")["Sales"].sum().reset_index(), x="Region", y="Sales", title="Total Sales by Region", text_auto=True)
st.plotly_chart(fig_sales_region, use_container_width=True)

# Sales vs. Profit Scatter
fig_scatter = px.scatter(df_filtered, x="Sales", y="Profit", color="Category", hover_data=["Sub-Category", "Customer Name"], title="Sales vs Profit")
st.plotly_chart(fig_scatter, use_container_width=True)

# Sales by Category
fig_sales_category = px.pie(df_filtered, names="Category", values="Sales", title="Sales Distribution by Category")
st.plotly_chart(fig_sales_category, use_container_width=True)

# Profit Trends
df_time = df_filtered.resample("M", on="Order Date").sum().reset_index()
fig_profit_trend = px.line(df_time, x="Order Date", y="Profit", title="Monthly Profit Trends")
st.plotly_chart(fig_profit_trend, use_container_width=True)

st.write("### Insights:")
st.markdown("- The sales distribution helps identify the highest-performing regions.")
st.markdown("- The return rate metric provides insights into potential product issues.")
st.markdown("- The scatter plot highlights high-profit vs low-profit transactions.")
