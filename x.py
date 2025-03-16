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

# Ensure selected KPI exists in the dataset
#valid_kpis = ["Sales", "Profit"]
valid_charts = ["Line Chart", "Bar Chart", "Pie Chart", "Scatter Plot"]
valid_x = orders_df.columns.tolist()
valid_y = orders_df.columns.tolist()

st.set_page_config(page_title="Superstore Dashboard", layout="wide")
# Apply custom dark theme style
st.markdown(
    """
    <style>
        body {
            background-color: black;
            color: white;
        }
        .css-1d391kg, .css-1v3fvcr, .css-18e3th9 {
            background-color: black !important;
            color: white !important;
        }
    </style>
    """,
    unsafe_allow_html=True
)
st.title("Superstore Dashboard")

# Sidebar Filters
st.sidebar.header("Filter Options")
selected_region = st.sidebar.multiselect("Select Region", options=orders_df["Region"].unique(), default=orders_df["Region"].unique())
selected_category = st.sidebar.multiselect("Select Category", options=orders_df["Category"].unique(), default=orders_df["Category"].unique())
#selected_kpi = st.sidebar.selectbox("Select KPI", valid_kpis)
selected_chart = st.sidebar.selectbox("Select Chart Type", valid_charts)
selected_x = st.sidebar.selectbox("Select X-axis", valid_x)
selected_y = st.sidebar.selectbox("Select Y-axis", valid_y)

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

# Dynamic KPI Chart
if selected_x in df_filtered.columns and selected_y in df_filtered.columns:
    if selected_chart == "Line Chart":
        fig_kpi = px.line(df_filtered, x=selected_x, y=selected_y, title=f"{selected_y} Trend Over {selected_x}")
    elif selected_chart == "Bar Chart":
        fig_kpi = px.bar(df_filtered, x=selected_x, y=selected_y, title=f"{selected_y} Over {selected_x}")
    elif selected_chart == "Pie Chart":
        fig_kpi = px.pie(df_filtered, names=selected_x, values=selected_y, title=f"{selected_y} Distribution by {selected_x}")
    elif selected_chart == "Scatter Plot":
        fig_kpi = px.scatter(df_filtered, x=selected_x, y=selected_y, title=f"{selected_y} Scatter Plot vs {selected_x}")
    
    st.plotly_chart(fig_kpi, use_container_width=True)

"""st.write("### Insights:")
st.markdown("- The sales distribution helps identify the highest-performing regions.")
st.markdown("- The return rate metric provides insights into potential product issues.")
st.markdown("- The scatter plot highlights high-profit vs low-profit transactions.")
"""