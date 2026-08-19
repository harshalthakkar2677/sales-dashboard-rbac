from sales_cac_utils import calculate_payouts
import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Cost Per Sale Dashboard", layout="wide")

def highlight_cps_category(val):
    if pd.isna(val):
        return ""
    val = str(val)
    if val == "Good":
        return "background-color: #d4edda; color: #155724; font-weight: bold;"
    elif val == "Needs Improvement":
        return "background-color: #fff3cd; color: #856404; font-weight: bold;"
    elif val == "Alarming":
        return "background-color: #f8d7da; color: #721c24; font-weight: bold;"
    return ""

def style_cps_table(df):
    styled = df.style
    if "CPS Category" in df.columns:
        try:
            styled = styled.map(highlight_cps_category, subset=["CPS Category"])
        except:
            styled = styled.applymap(highlight_cps_category, subset=["CPS Category"])

    fmt_cols = {}
    for col in ["Fixed_CTC", "Variable_Payout", "Pending_Payout", "Total_Payout", "CAC", "Avg CPS"]:
        if col in df.columns:
            fmt_cols[col] = "{:,.0f}"
    styled = styled.format(fmt_cols)
    return styled

def style_audit_table(df):
    styled = df.style
    fmt_cols = {}
    for col in ["Plan Value", "ROW_PAYOUT", "ROW_PENDING"]:
        if col in df.columns:
            fmt_cols[col] = "{:,.0f}"
    return styled.format(fmt_cols)

def classify_cps(cac):
    if pd.isna(cac):
        return "Missing"
    if cac < 1800:
        return "Good"
    elif cac <= 2200:
        return "Needs Improvement"
    return "Alarming"

def incentive_status(variable_payout):
    if pd.isna(variable_payout) or variable_payout <= 0:
        return "No Incentive"
    return "Incentive Earned"

def add_total_row(df, label_col=None, label="Total"):
    out = df.copy()
    numeric_cols = out.select_dtypes(include=["number"]).columns.tolist()

    total_dict = {}
    for col in out.columns:
        if col in numeric_cols:
            total_dict[col] = out[col].sum()
        else:
            total_dict[col] = ""

    if label_col and label_col in out.columns:
        total_dict[label_col] = label

    total_df = pd.DataFrame([total_dict])
    return pd.concat([out, total_df], ignore_index=True)

st.title("📊 Employee-wise Cost Per Sale (CPS / CAC)")
st.markdown(
    """
    **Rule applied:** CPS is calculated only for employees who are:
    - in **Sales** department
    - **Active**
    - present in the **CTC file**
    """
)

sales_file = st.file_uploader("Upload New Registration Report", type=["csv"])
wb_file = st.file_uploader("Upload New Winback Report", type=["csv"])
ctc_file = st.file_uploader("Upload Active Sales CTC File", type=["csv"])

if sales_file and wb_file and ctc_file:
    monthly_cac, detail_rows = calculate_payouts(
        sales_file=sales_file,
        wb_file=wb_file,
        ctc_file=ctc_file
    )

    monthly_cac["CPS Category"] = monthly_cac["CAC"].apply(classify_cps)
    monthly_cac["Incentive Status"] = monthly_cac["Variable_Payout"].apply(incentive_status)

    display_df = monthly_cac.copy()
    detail_df = detail_rows.copy()

    display_df["MonthSort"] = pd.to_datetime(display_df["MonthYear"], format="%b-%Y", errors="coerce")
    month_options = ["All"] + display_df.sort_values("MonthSort")["MonthYear"].dropna().drop_duplicates().tolist()
    city_options = ["All"] + sorted(display_df["City"].dropna().astype(str).unique().tolist())
    scheme_options = ["All"] + sorted(display_df["SCHEME"].dropna().astype(str).unique().tolist())

    c1, c2, c3 = st.columns(3)
    with c1:
        selected_month = st.selectbox("Select Month", month_options)
    with c2:
        selected_city = st.selectbox("Select City", city_options)
    with c3:
        selected_scheme = st.selectbox("Select Scheme", scheme_options)

    filtered = display_df.copy()
    filtered_detail = detail_df.copy()

    if selected_month != "All":
        filtered = filtered[filtered["MonthYear"] == selected_month]
        filtered_detail = filtered_detail[filtered_detail["MonthYear"] == selected_month]

    if selected_city != "All":
        filtered = filtered[filtered["City"] == selected_city]
        filtered_detail = filtered_detail[filtered_detail["City"] == selected_city]

    if selected_scheme != "All":
        filtered = filtered[filtered["SCHEME"] == selected_scheme]
        filtered_detail = filtered_detail[filtered_detail["SCHEME"] == selected_scheme]

    total_execs = filtered["EMP Code"].nunique()
    total_acts = filtered["Total_Activations"].sum()
    total_cost = filtered["Total_Payout"].sum()
    avg_cac = round(total_cost / total_acts, 0) if total_acts else 0
    no_incentive_execs = filtered[filtered["Variable_Payout"] <= 0]["EMP Code"].nunique()

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Executives", int(total_execs))
    k2.metric("Total Activations", int(total_acts))
    k3.metric("Total Cost", f"₹{total_cost:,.0f}")
    k4.metric("Average CPS / CAC", f"₹{avg_cac:,.0f}")
    k5.metric("Executives with No Incentive", int(no_incentive_execs))

    st.subheader("CPS Category Summary")
    category_emp = filtered[["EMP Code", "CPS Category"]].drop_duplicates()

    category_summary = (
        category_emp.groupby("CPS Category", as_index=False)
        .agg(Employees=("EMP Code", "nunique"))
    )

    category_value = (
        filtered.groupby("CPS Category", as_index=False)
        .agg(
            Total_Activations=("Total_Activations", "sum"),
            Total_Payout=("Total_Payout", "sum")
        )
    )

    category_summary = category_summary.merge(category_value, on="CPS Category", how="left")
    category_summary["Avg CPS"] = (
        category_summary["Total_Payout"] / category_summary["Total_Activations"].replace(0, pd.NA)
    ).fillna(0).round(0)

    st.dataframe(style_cps_table(category_summary), use_container_width=True)

    st.subheader("Incentive Qualification Summary")
    incentive_emp = filtered[["EMP Code", "Incentive Status"]].drop_duplicates()

    incentive_summary = (
        incentive_emp.groupby("Incentive Status", as_index=False)
        .agg(Employees=("EMP Code", "nunique"))
    )

    incentive_value = (
        filtered.groupby("Incentive Status", as_index=False)
        .agg(
            Total_Activations=("Total_Activations", "sum"),
            Total_Payout=("Total_Payout", "sum")
        )
    )

    incentive_summary = incentive_summary.merge(incentive_value, on="Incentive Status", how="left")
    st.dataframe(incentive_summary, use_container_width=True)

    st.subheader("Month-on-Month CPS / CAC Trend")
    trend_source = filtered.copy()

    trend = trend_source.groupby("MonthYear", as_index=False).agg(
        Total_Payout=("Total_Payout", "sum"),
        Total_Activations=("Total_Activations", "sum")
    )
    trend["MonthSort"] = pd.to_datetime(trend["MonthYear"], format="%b-%Y", errors="coerce")
    trend = trend.sort_values("MonthSort")
    trend["CAC"] = (trend["Total_Payout"] / trend["Total_Activations"].replace(0, pd.NA)).fillna(0).round(0)

    if not trend.empty:
        fig_trend = px.line(trend, x="MonthYear", y="CAC", markers=True, title="Month-on-Month CPS / CAC")
        st.plotly_chart(fig_trend, use_container_width=True)

    st.subheader("Employee-wise CPS Summary")
    show_cols = [
        "EMP Code", "Name", "City", "MonthYear",
        "Installs", "Winbacks", "Total_Activations",
        "SCHEME", "Fixed_CTC", "Variable_Payout",
        "Pending_Payout", "Total_Payout", "CAC",
        "CPS Category", "Incentive Status"
    ]
    show_cols = [c for c in show_cols if c in filtered.columns]
    summary_table = add_total_row(filtered[show_cols].copy(), label_col="Name")
    st.dataframe(style_cps_table(summary_table), use_container_width=True)

    with st.expander("View Activation-level CPS Audit Detail"):
        detail_cols = [
            "EMP Code", "Name", "MonthYear", "Source", "City",
            "Plan Value", "SPEED (Mbps)", "VALIDITY In Months",
            "Installs", "Winbacks", "Total Activations",
            "SCHEME", "ROW_PAYOUT", "ROW_PENDING", "ROW_REMARK", "DEBUG_MATCH"
        ]
        detail_cols = [c for c in detail_cols if c in filtered_detail.columns]

        st.dataframe(style_audit_table(filtered_detail[detail_cols]), use_container_width=True)

        detail_download = filtered_detail[detail_cols].to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Download Filtered Audit Detail",
            data=detail_download,
            file_name="filtered_employee_activation_level_cps_detail.csv",
            mime="text/csv"
        )

    csv_summary = filtered.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Download CPS Summary",
        data=csv_summary,
        file_name="employee_monthly_cps_summary.csv",
        mime="text/csv"
    )

else:
    st.info("Please upload all 3 files to calculate employee-wise CPS.")