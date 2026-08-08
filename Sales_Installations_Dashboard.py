import os   
import streamlit as st
import pandas as pd
import plotly.express as px

# --- Logo and Header ---
st.image("company_logo.png", width=200)
st.markdown("## Welcome to the Sales & Installation Dashboard")
st.markdown("🔐 Sales Performance & Installation SLA Login")

# --- Load Access Master ---
access_master = pd.read_excel("Access Master.xlsx")

# --- Login Form ---
username = st.text_input("Username")
password = st.text_input("Password", type="password")

if st.button("Login"):
    # Check if user exists in Access Master
    user_row = access_master[
        (access_master["Username"] == username) &
        (access_master["Password"] == password)
    ]

    if not user_row.empty:
        role = user_row["Role"].values[0]
        region = user_row["Region/City"].values[0]

        # --- Load Dashboard Data ---
        DATA_FILE = "New Registration Report.csv"
        WINBACK_FILE = "New Winback Report.csv"
        df = pd.read_csv(DATA_FILE)
        wb = pd.read_csv(WINBACK_FILE)

        # --- Role-based Dashboards ---
        if role == "Manager":
            st.header("📊 Manager Dashboard - All Cities")
            st.dataframe(df)

        elif role == "Regional Lead":
            # Split comma-separated cities into a list
            allowed_cities = [city.strip() for city in region.split(",")]
            st.header(f"🌍 Regional Dashboard - {', '.join(allowed_cities)}")
            st.write("Access limited to city-level performance and SLA metrics.")

            regional_df = df[df['City'].isin(allowed_cities)]
            st.dataframe(regional_df)

            regional_wb = wb[wb['City'].isin(allowed_cities)]
            st.dataframe(regional_wb)

        elif role == "CEO":
            st.header("🏢 Sales & Installations Dashboard - CEO & Sales Team View")
            st.dataframe(df)
            st.dataframe(wb)

        else:
            st.warning("Role not recognized.")

    else:
        st.error("Invalid username or password")

    # --------------------------------------------------
    # Page Config
    # --------------------------------------------------
    st.set_page_config(page_title="Sales & Installations Dashboard", layout="wide")

    # --------------------------------------------------
    # Header
    # --------------------------------------------------
    logo_path = os.path.join(os.path.dirname(__file__), "company_logo.png")

    col1, col2 = st.columns([1, 5])
    with col1:
        if os.path.exists(logo_path):
            st.image(logo_path, width=120)
    with col2:
        st.title("Sales & Installations Dashboard")
        st.markdown("#### CEO & Sales Team View")

    # --------------------------------------------------
    # Refresh Control
    # --------------------------------------------------
    DATA_FILE = r"C:\Users\Harshal Thakkar\Dashboard\Sales & Installation\sales-dashboard-rbac\New Registration Report.csv"
    WINBACK_FILE = r"C:\Users\Harshal Thakkar\Dashboard\Sales & Installation\sales-dashboard-rbac\New Winback Report.csv"

    refresh_col1, refresh_col2 = st.columns([1, 5])
    with refresh_col1:
        if st.button("🔄 Refresh Data"):
            st.cache_data.clear()
            st.success("Data cache cleared. Latest file will be loaded.")

    file_mtime = os.path.getmtime(DATA_FILE) if os.path.exists(DATA_FILE) else 0

    # --------------------------------------------------
    # Helpers
    # --------------------------------------------------
    def format_value(num):
        if pd.isna(num):
            return "0"
        if num >= 1_000_000:
            return f"{num/1_000_000:.1f} M"
        elif num >= 100_000:
            return f"{num/100_000:.1f} L"
        else:
            return f"{round(num):,}"

    def format_currency_compact(num):
        if pd.isna(num):
            return "₹0"
        if num >= 1_000_000:
            return f"₹{num/1_000_000:.2f}M"
        elif num >= 100_000:
            return f"₹{num/100_000:.2f}L"
        elif num >= 1_000:
            return f"₹{num/1000:.0f}K"
        else:
            return f"₹{num:,.0f}"

    def sla_bucket(x):
        if pd.isna(x):
            return "Missing"
        elif x <= 2:
            return "0-2 days"
        elif x <= 5:
            return "2-5 days"
        elif x <= 7:
            return "5-7 days"
        else:
            return "Above 7 days"

    def generate_summary(dataframe, level="Executive", group_col="Exec Name"):
        if dataframe.empty or group_col not in dataframe.columns:
            return "No summary available for selected filters."
        summary = dataframe.groupby(group_col)["ACCOUNT NO"].count().sort_values(ascending=False)
        if summary.empty:
            return "No summary available for selected filters."
        top_name = summary.index[0]
        top_val = int(summary.iloc[0])
        return f"Top {level}: **{top_name}** with **{top_val} contributions** for the selected filters."

    def arpu_bucket(x):
        if pd.isna(x):
            return "Missing"
        elif x <= 300:
            return "upto 300"
        elif x <= 500:
            return "301-500"
        elif x <= 750:
            return "501-750"
        else:
            return "751+"

    def normalize_exec_status(x):
        if pd.isna(x):
            return "Missing"
        x = str(x).strip().upper()
        mapping = {
            "ACTIVE": "Active",
            "INACTIVE": "Inactive",
            "EXIT": "Inactive",
            "RESIGNED": "Inactive",
            "TERMINATED": "Inactive",
            "SUSPEND": "Inactive",
            "SUSPENDED": "Inactive"
        }
        return mapping.get(x, x.title())

    def normalize_department(x):
        if pd.isna(x):
            return "Missing"
        x = str(x).strip().lower()
        mapping = {
            "channel partner": "Channel Partner",
            "channelpartner": "Channel Partner",
            "channel-partner": "Channel Partner",
            "channel partner ": "Channel Partner",
            "sales direct": "Sales Direct",
            "esg - sales": "ESG - Sales",
            "online": "Online",
            "retention": "Retention",
            "technical": "Technical",
            "sales": "Sales"
        }
        return mapping.get(x, str(x).strip().title())

    def ageing_bucket(months):
        if pd.isna(months):
            return "Missing"
        elif months <= 3:
            return "Less Than 3 Months"
        elif months <= 12:
            return "4-12 Months"
        else:
            return "Above 12 Months"

    def target_by_ageing(bucket):
        mapping = {
            "Less Than 3 Months": 5,
            "4-12 Months": 8,
            "Above 12 Months": 12
        }
        return mapping.get(bucket, 0)

    def expectation_label(avg_sales, target):
        if pd.isna(avg_sales) or pd.isna(target):
            return "Missing"
        if avg_sales < target:
            return "Below Expectations"
        elif avg_sales == target:
            return "Meeting Expectations"
        else:
            return "Exceptional"

    def sort_month_df(df_in, month_col="MonthYear"):
        out = df_in.copy()
        if month_col in out.columns:
            out[f"{month_col}_dt"] = pd.to_datetime(out[month_col], format="%b-%Y", errors="coerce")
            out = out.sort_values(f"{month_col}_dt")
        return out

    # --------------------------------------------------
    # Cached Data Loader
    # --------------------------------------------------
    @st.cache_data
    def load_data(file_mtime):
        # -----------------------------
        # Load Installation Data
        # -----------------------------
        df = pd.read_csv(DATA_FILE)
        df.columns = df.columns.str.strip()

        if "CREATION DATE" in df.columns:
            df["CREATION DATE"] = pd.to_datetime(df["CREATION DATE"], dayfirst=True, errors="coerce")

        if "INSTALLATION DATE" in df.columns:
            raw_install = pd.read_csv(DATA_FILE, usecols=["INSTALLATION DATE"])["INSTALLATION DATE"].astype(str).str.strip()
            parsed_1 = pd.to_datetime(raw_install, format="%d-%m-%Y", errors="coerce")
            parsed_2 = pd.to_datetime(raw_install, format="%d/%m/%Y", errors="coerce")
            parsed_3 = pd.to_datetime(raw_install, format="%Y-%m-%d", errors="coerce")
            parsed_4 = pd.to_datetime(raw_install, dayfirst=True, errors="coerce")
            df["INSTALLATION DATE"] = parsed_1.fillna(parsed_2).fillna(parsed_3).fillna(parsed_4)

            df["MonthStart"] = df["INSTALLATION DATE"].dt.to_period("M").dt.to_timestamp()
            df["MonthYear"] = df["MonthStart"].dt.strftime("%b-%Y")
            df["Year"] = df["INSTALLATION DATE"].dt.year
        else:
            df["MonthStart"] = pd.NaT
            df["MonthYear"] = None
            df["Year"] = None

        if "Plan Value" in df.columns:
            df["Plan Value"] = pd.to_numeric(df["Plan Value"], errors="coerce").fillna(0)
        else:
            df["Plan Value"] = 0

        if "ARPU" in df.columns:
            df["ARPU"] = pd.to_numeric(df["ARPU"], errors="coerce")
            df["ARPU_BUCKET"] = df["ARPU"].apply(arpu_bucket)
        else:
            df["ARPU"] = None
            df["ARPU_BUCKET"] = "Missing"

        if "ACCOUNT NO" not in df.columns:
            df["ACCOUNT NO"] = None

        if "CREATION DATE" in df.columns and "INSTALLATION DATE" in df.columns:
            df["TAT_days"] = (df["INSTALLATION DATE"] - df["CREATION DATE"]).dt.days
            df["SLA_Category"] = df["TAT_days"].apply(sla_bucket)
        else:
            df["TAT_days"] = None
            df["SLA_Category"] = "Missing"

        df["Record_Type"] = "Installation"

        # -----------------------------
        # Normalize Installation Data
        # -----------------------------
        if "SALES CODE" in df.columns:
            df["SALES CODE"] = df["SALES CODE"].astype(str).str.strip()

        if "SALES EXEC NAME" in df.columns:
            df["SALES EXEC NAME"] = df["SALES EXEC NAME"].astype(str).str.strip()

        if "SALES EXEC STATUS" in df.columns:
            df["SALES EXEC STATUS"] = df["SALES EXEC STATUS"].astype(str).str.strip()
            df["EXEC_STATUS_CLEAN"] = df["SALES EXEC STATUS"].apply(normalize_exec_status)
        else:
            df["EXEC_STATUS_CLEAN"] = "Missing"

        if "DEPARTMENT" in df.columns:
            df["DEPARTMENT"] = df["DEPARTMENT"].astype(str).str.strip()
            df["DEPARTMENT_CLEAN"] = df["DEPARTMENT"].apply(normalize_department)
        else:
            df["DEPARTMENT_CLEAN"] = "Missing"

        if "SALES EXEC DOJ" in df.columns:
            raw_doj = pd.read_csv(DATA_FILE, usecols=["SALES EXEC DOJ"])["SALES EXEC DOJ"].astype(str).str.strip()
            d1 = pd.to_datetime(raw_doj, format="%d-%b-%y", errors="coerce")
            d2 = pd.to_datetime(raw_doj, format="%d-%b-%Y", errors="coerce")
            d3 = pd.to_datetime(raw_doj, dayfirst=True, errors="coerce")
            df["SALES EXEC DOJ"] = d1.fillna(d2).fillna(d3)
        else:
            df["SALES EXEC DOJ"] = pd.NaT

        # -----------------------------
        # Employee Master from Installations
        # -----------------------------
        if "SALES CODE" in df.columns:
            sort_col = "INSTALLATION DATE" if "INSTALLATION DATE" in df.columns else "CREATION DATE"

            emp_master = (
                df.sort_values(sort_col)
                  .dropna(subset=["SALES CODE"])
                  .groupby("SALES CODE", as_index=False)
                  .last()[["SALES CODE", "SALES EXEC NAME", "EXEC_STATUS_CLEAN", "DEPARTMENT_CLEAN", "SALES EXEC DOJ"]]
                  .rename(columns={
                      "SALES EXEC NAME": "EXEC_NAME_MASTER",
                      "EXEC_STATUS_CLEAN": "EXEC_STATUS_MASTER",
                      "DEPARTMENT_CLEAN": "DEPARTMENT_MASTER",
                      "SALES EXEC DOJ": "DOJ_MASTER"
                  })
            )

            df = df.drop(columns=[c for c in ["EXEC_NAME_MASTER", "EXEC_STATUS_MASTER", "DEPARTMENT_MASTER", "DOJ_MASTER"] if c in df.columns], errors="ignore")
            df = df.merge(emp_master, on="SALES CODE", how="left")

            df["EXEC_NAME_FINAL"] = df["EXEC_NAME_MASTER"].fillna(df["SALES EXEC NAME"])
            df["EXEC_STATUS_FINAL"] = df["EXEC_STATUS_MASTER"].fillna(df["EXEC_STATUS_CLEAN"])
            df["DEPARTMENT_FINAL"] = df["DEPARTMENT_MASTER"].fillna(df["DEPARTMENT_CLEAN"])
            df["SALES EXEC DOJ FINAL"] = df["DOJ_MASTER"]
        else:
            df["EXEC_NAME_FINAL"] = df["SALES EXEC NAME"]
            df["EXEC_STATUS_FINAL"] = df["EXEC_STATUS_CLEAN"]
            df["DEPARTMENT_FINAL"] = df["DEPARTMENT_CLEAN"]
            df["SALES EXEC DOJ FINAL"] = df["SALES EXEC DOJ"]

        today = pd.Timestamp.today().normalize()
        df["Exec_Ageing_Months"] = ((today - df["SALES EXEC DOJ FINAL"]).dt.days / 30.44).round(1)
        df["Exec_Ageing_Bucket"] = df["Exec_Ageing_Months"].apply(ageing_bucket)
        df["Target_Avg_Month"] = df["Exec_Ageing_Bucket"].apply(target_by_ageing)

        # -----------------------------
        # Load Winback Data
        # -----------------------------
        if os.path.exists(WINBACK_FILE):
            wb = pd.read_csv(WINBACK_FILE)
            wb.columns = wb.columns.str.strip()

            if "INSTALLATION DATE" in wb.columns:
                raw_wb_install = pd.read_csv(WINBACK_FILE, usecols=["INSTALLATION DATE"])["INSTALLATION DATE"].astype(str).str.strip()
                w1 = pd.to_datetime(raw_wb_install, format="%d-%m-%y %H:%M", errors="coerce")
                w2 = pd.to_datetime(raw_wb_install, format="%d-%m-%Y", errors="coerce")
                w3 = pd.to_datetime(raw_wb_install, format="%d/%m/%Y", errors="coerce")
                w4 = pd.to_datetime(raw_wb_install, dayfirst=True, errors="coerce")
                wb["INSTALLATION DATE"] = w1.fillna(w2).fillna(w3).fillna(w4)

                wb["MonthStart"] = wb["INSTALLATION DATE"].dt.to_period("M").dt.to_timestamp()
                wb["MonthYear"] = wb["MonthStart"].dt.strftime("%b-%Y")
                wb["Year"] = wb["INSTALLATION DATE"].dt.year
            else:
                wb["MonthStart"] = pd.NaT
                wb["MonthYear"] = None
                wb["Year"] = None

            if "Plan Value" in wb.columns:
                wb["Plan Value"] = pd.to_numeric(wb["Plan Value"], errors="coerce").fillna(0)
            else:
                wb["Plan Value"] = 0

            if "ARPU" in wb.columns:
                wb["ARPU"] = pd.to_numeric(wb["ARPU"], errors="coerce")
                wb["ARPU_BUCKET"] = wb["ARPU"].apply(arpu_bucket)
            else:
                wb["ARPU"] = None
                wb["ARPU_BUCKET"] = "Missing"

            if "SALES CODE" in wb.columns:
                wb["SALES CODE"] = wb["SALES CODE"].astype(str).str.strip()

            if "SALES EXEC NAME" in wb.columns:
                wb["SALES EXEC NAME"] = wb["SALES EXEC NAME"].astype(str).str.strip()

            if "SALES EXEC STATUS" in wb.columns:
                wb["SALES EXEC STATUS"] = wb["SALES EXEC STATUS"].astype(str).str.strip()
                wb["EXEC_STATUS_CLEAN"] = wb["SALES EXEC STATUS"].apply(normalize_exec_status)
            else:
                wb["EXEC_STATUS_CLEAN"] = "Missing"

            if "DEPARTMENT" in wb.columns:
                wb["DEPARTMENT"] = wb["DEPARTMENT"].astype(str).str.strip()
                wb["DEPARTMENT_CLEAN"] = wb["DEPARTMENT"].apply(normalize_department)
            else:
                wb["DEPARTMENT_CLEAN"] = "Missing"

            if "SALES EXEC DOJ" in wb.columns:
                raw_wb_doj = pd.read_csv(WINBACK_FILE, usecols=["SALES EXEC DOJ"])["SALES EXEC DOJ"].astype(str).str.strip()
                wd1 = pd.to_datetime(raw_wb_doj, format="%d-%b-%y", errors="coerce")
                wd2 = pd.to_datetime(raw_wb_doj, format="%d-%b-%Y", errors="coerce")
                wd3 = pd.to_datetime(raw_wb_doj, dayfirst=True, errors="coerce")
                wb["SALES EXEC DOJ"] = wd1.fillna(wd2).fillna(wd3)
            else:
                wb["SALES EXEC DOJ"] = pd.NaT

            wb["Record_Type"] = "Winback"

            if "SALES CODE" in wb.columns and "SALES CODE" in emp_master.columns:
                wb = wb.merge(emp_master, on="SALES CODE", how="left")
                wb["EXEC_NAME_FINAL"] = wb["EXEC_NAME_MASTER"].fillna(wb["SALES EXEC NAME"])
                wb["EXEC_STATUS_FINAL"] = wb["EXEC_STATUS_MASTER"].fillna(wb["EXEC_STATUS_CLEAN"])
                wb["DEPARTMENT_FINAL"] = wb["DEPARTMENT_MASTER"].fillna(wb["DEPARTMENT_CLEAN"])
                wb["SALES EXEC DOJ FINAL"] = wb["DOJ_MASTER"]
            else:
                wb["EXEC_NAME_FINAL"] = wb["SALES EXEC NAME"]
                wb["EXEC_STATUS_FINAL"] = wb["EXEC_STATUS_CLEAN"]
                wb["DEPARTMENT_FINAL"] = wb["DEPARTMENT_CLEAN"]
                wb["SALES EXEC DOJ FINAL"] = wb["SALES EXEC DOJ"]

            wb["Exec_Ageing_Months"] = ((today - wb["SALES EXEC DOJ FINAL"]).dt.days / 30.44).round(1)
            wb["Exec_Ageing_Bucket"] = wb["Exec_Ageing_Months"].apply(ageing_bucket)
            wb["Target_Avg_Month"] = wb["Exec_Ageing_Bucket"].apply(target_by_ageing)
        else:
            wb = pd.DataFrame()

        return df, wb

    df, wb = load_data(file_mtime)

    # --------------------------------------------------
    # Contribution Mode Filter
    # --------------------------------------------------
    contribution_mode = st.sidebar.selectbox(
        "Contribution Mode",
        ["Installations + Winback", "Installation", "Winback"],
        key="contribution_mode"
    )

    # --------------------------------------------------
    # Sidebar Filters
    # --------------------------------------------------
    st.sidebar.header("Filters")

    city = st.sidebar.selectbox(
        "City",
        ["All"] + sorted(df["City"].dropna().astype(str).unique()) if "City" in df.columns else ["All"]
    )

    department = st.sidebar.selectbox(
        "Department / Channel",
        ["All"] + sorted(df["DEPARTMENT_FINAL"].dropna().astype(str).unique()) if "DEPARTMENT_FINAL" in df.columns else ["All"]
    )

    exec_status = st.sidebar.selectbox(
        "Executive Status",
        ["All"] + sorted(df["EXEC_STATUS_FINAL"].dropna().astype(str).unique()) if "EXEC_STATUS_FINAL" in df.columns else ["All"]
    )

    ageing_bucket_filter = st.sidebar.selectbox(
        "Executive Ageing",
        ["All", "Less Than 3 Months", "4-12 Months", "Above 12 Months"]
    )

    if contribution_mode == "Winback":
        source_df = wb.copy()
    elif contribution_mode == "Installation":
        source_df = df.copy()
    else:
        source_df = pd.concat([df.copy(), wb.copy()], ignore_index=True) if not wb.empty else df.copy()

    if "EXEC_NAME_FINAL" in source_df.columns:
        exec_df = source_df.copy()
        if city != "All" and "City" in exec_df.columns:
            exec_df = exec_df[exec_df["City"] == city]
        if department != "All" and "DEPARTMENT_FINAL" in exec_df.columns:
            exec_df = exec_df[exec_df["DEPARTMENT_FINAL"] == department]
        if exec_status != "All" and "EXEC_STATUS_FINAL" in exec_df.columns:
            exec_df = exec_df[exec_df["EXEC_STATUS_FINAL"] == exec_status]
        if ageing_bucket_filter != "All" and "Exec_Ageing_Bucket" in exec_df.columns:
            exec_df = exec_df[exec_df["Exec_Ageing_Bucket"] == ageing_bucket_filter]

        execu_list = sorted(exec_df["EXEC_NAME_FINAL"].dropna().astype(str).unique())
    else:
        execu_list = []

    execu = st.sidebar.selectbox("Executive", ["All"] + execu_list)

    if "MonthYear" in source_df.columns:
        month_options = source_df["MonthYear"].dropna().unique().tolist()
        month_options = sorted(month_options, key=lambda x: pd.to_datetime(x, format="%b-%Y"))
    else:
        month_options = []

    month = st.sidebar.selectbox("Month", ["All"] + month_options)

    speed = st.sidebar.selectbox(
        "Speed (Mbps)",
        ["All"] + sorted(source_df["SPEED (Mbps)"].dropna().unique()) if "SPEED (Mbps)" in source_df.columns else ["All"]
    )

    validity = st.sidebar.selectbox(
        "Validity (Months)",
        ["All"] + sorted(source_df["VALIDITY In Months"].dropna().unique()) if "VALIDITY In Months" in source_df.columns else ["All"]
    )

    network_type = st.sidebar.selectbox(
        "Network Type",
        ["All"] + sorted(source_df["NETWORK TYPE"].dropna().astype(str).unique()) if "NETWORK TYPE" in source_df.columns else ["All"]
    )

    # --------------------------------------------------
    # Apply filters
    # --------------------------------------------------
    filtered_df_base = source_df.copy()

    if city != "All" and "City" in filtered_df_base.columns:
        filtered_df_base = filtered_df_base[filtered_df_base["City"] == city]
    if department != "All" and "DEPARTMENT_FINAL" in filtered_df_base.columns:
        filtered_df_base = filtered_df_base[filtered_df_base["DEPARTMENT_FINAL"] == department]
    if exec_status != "All" and "EXEC_STATUS_FINAL" in filtered_df_base.columns:
        filtered_df_base = filtered_df_base[filtered_df_base["EXEC_STATUS_FINAL"] == exec_status]
    if ageing_bucket_filter != "All" and "Exec_Ageing_Bucket" in filtered_df_base.columns:
        filtered_df_base = filtered_df_base[filtered_df_base["Exec_Ageing_Bucket"] == ageing_bucket_filter]
    if execu != "All" and "EXEC_NAME_FINAL" in filtered_df_base.columns:
        filtered_df_base = filtered_df_base[filtered_df_base["EXEC_NAME_FINAL"] == execu]
    if speed != "All" and "SPEED (Mbps)" in filtered_df_base.columns:
        filtered_df_base = filtered_df_base[filtered_df_base["SPEED (Mbps)"] == speed]
    if validity != "All" and "VALIDITY In Months" in filtered_df_base.columns:
        filtered_df_base = filtered_df_base[filtered_df_base["VALIDITY In Months"] == validity]
    if network_type != "All" and "NETWORK TYPE" in filtered_df_base.columns:
        filtered_df_base = filtered_df_base[filtered_df_base["NETWORK TYPE"] == network_type]

    filtered_df = filtered_df_base.copy()
    if month != "All" and "MonthYear" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["MonthYear"] == month]

    # --------------------------------------------------
    # Tabs
    # --------------------------------------------------
    tab1, tab2, tab3 = st.tabs(["SLA & Executive Performance", "Customer Profiling", "Sales Head Analytics"])
    ## Part 2 of 4

    # ==================================================
    # TAB 1: Operations & Executive
    # ==================================================
    with tab1:
        # Executive Summary for selected executive
        if execu != "All" and "EXEC_NAME_FINAL" in filtered_df_base.columns:
            exec_data = filtered_df_base[filtered_df_base["EXEC_NAME_FINAL"] == execu]
            total_count = exec_data["ACCOUNT NO"].count()
            total_value = exec_data["Plan Value"].sum()
            months_available = exec_data["MonthYear"].nunique()
            avg_per_month_count = exec_data.groupby("MonthYear")["ACCOUNT NO"].count().mean() if not exec_data.empty else 0
            avg_per_month_value = exec_data.groupby("MonthYear")["Plan Value"].sum().mean() if not exec_data.empty else 0
            per_sales_value = total_value / total_count if total_count > 0 else 0

            st.subheader(f"Executive Dashboard: {execu}")
            k1, k2, k3, k4, k5 = st.columns(5)
            k1.metric(f"Total Count ({months_available} months)", int(total_count))
            k2.metric("Total Value", f"{total_value:,.0f}")
            k3.metric("Average Count per Month", f"{0 if pd.isna(avg_per_month_count) else avg_per_month_count:.0f}")
            k4.metric("Average Value per Month", f"{0 if pd.isna(avg_per_month_value) else avg_per_month_value:.0f}")
            k5.metric("Per-Sales Value", f"{per_sales_value:,.0f}")

        # MoM Executive Contribution
        if city != "All" and "EXEC_NAME_FINAL" in filtered_df_base.columns and "MonthStart" in filtered_df_base.columns:
            st.subheader("Month-on-Month Executive Contribution")

            if contribution_mode == "Installations + Winback":
                mom_exec = (
                    filtered_df_base.groupby(["EXEC_NAME_FINAL", "MonthStart", "Record_Type"])
                    .agg(
                        Count=("ACCOUNT NO", "count"),
                        Value=("Plan Value", "sum")
                    )
                    .reset_index()
                )

                if not mom_exec.empty:
                    mom_exec["MonthYear"] = mom_exec["MonthStart"].dt.strftime("%b-%Y")

                    mom_pivot = mom_exec.pivot_table(
                        index=["EXEC_NAME_FINAL", "MonthStart", "MonthYear"],
                        columns="Record_Type",
                        values="Count",
                        fill_value=0
                    ).reset_index()

                    if "Installation" not in mom_pivot.columns:
                        mom_pivot["Installation"] = 0
                    if "Winback" not in mom_pivot.columns:
                        mom_pivot["Winback"] = 0

                    mom_pivot["Total"] = mom_pivot["Installation"] + mom_pivot["Winback"]
                    mom_pivot = mom_pivot.sort_values(["EXEC_NAME_FINAL", "MonthStart"])

                    mom_pivot["Prev_Total"] = mom_pivot.groupby("EXEC_NAME_FINAL")["Total"].shift(1)
                    mom_pivot["MoM_Total_%"] = ((mom_pivot["Total"] - mom_pivot["Prev_Total"]) / mom_pivot["Prev_Total"] * 100).round(1)

                    if execu != "All":
                        mom_pivot = mom_pivot[mom_pivot["EXEC_NAME_FINAL"] == execu].copy()

                    mom_pivot["MoM_Total_%"] = mom_pivot["MoM_Total_%"].apply(lambda x: f"{x}%" if pd.notna(x) else "")

                    st.dataframe(
                        mom_pivot[
                            ["EXEC_NAME_FINAL", "MonthYear", "Installation", "Winback", "Total", "Prev_Total", "MoM_Total_%"]
                        ].reset_index(drop=True),
                        use_container_width=True
                    )

                    mom_download = mom_pivot.copy()
                    csv_data = mom_download.to_csv(index=False).encode("utf-8")
                    st.download_button(
                        "📥 Download MoM Executive Contribution (CSV)",
                        csv_data,
                        file_name="MoM_Executive_Contribution.csv",
                        mime="text/csv"
                    )

            else:
                mom_exec = (
                    filtered_df_base.groupby(["EXEC_NAME_FINAL", "MonthStart"])
                    .agg(
                        Count=("ACCOUNT NO", "count"),
                        Value=("Plan Value", "sum")
                    )
                    .reset_index()
                    .sort_values(["EXEC_NAME_FINAL", "MonthStart"])
                )

                if not mom_exec.empty:
                    mom_exec["Prev_Count"] = mom_exec.groupby("EXEC_NAME_FINAL")["Count"].shift(1)
                    mom_exec["Prev_Value"] = mom_exec.groupby("EXEC_NAME_FINAL")["Value"].shift(1)
                    mom_exec["MoM_Count_%"] = ((mom_exec["Count"] - mom_exec["Prev_Count"]) / mom_exec["Prev_Count"] * 100).round(1)
                    mom_exec["MoM_Value_%"] = ((mom_exec["Value"] - mom_exec["Prev_Value"]) / mom_exec["Prev_Value"] * 100).round(1)
                    mom_exec["MonthYear"] = mom_exec["MonthStart"].dt.strftime("%b-%Y")

                    if execu != "All":
                        mom_exec = mom_exec[mom_exec["EXEC_NAME_FINAL"] == execu].copy()

                    mom_exec["MoM_Count_%"] = mom_exec["MoM_Count_%"].apply(lambda x: f"{x}%" if pd.notna(x) else "")
                    mom_exec["MoM_Value_%"] = mom_exec["MoM_Value_%"].apply(lambda x: f"{x}%" if pd.notna(x) else "")

                    st.dataframe(
                        mom_exec[
                            ["EXEC_NAME_FINAL", "MonthYear", "Count", "Value", "Prev_Count", "Prev_Value", "MoM_Count_%", "MoM_Value_%"]
                        ].reset_index(drop=True),
                        use_container_width=True
                    )

                    csv_data = mom_exec.to_csv(index=False).encode("utf-8")
                    st.download_button(
                        "📥 Download MoM Executive Performance (CSV)",
                        csv_data,
                        file_name="MoM_Executive_Performance.csv",
                        mime="text/csv"
                    )

        # SLA only for non-Winback mode
        if contribution_mode != "Winback":
            st.subheader("SLA Distribution")
            sla_chart = filtered_df.groupby("SLA_Category")["ACCOUNT NO"].count().reset_index()
            sla_chart.rename(columns={"ACCOUNT NO": "Count"}, inplace=True)

            if not sla_chart.empty and sla_chart["Count"].sum() > 0:
                sla_chart["Percent"] = (sla_chart["Count"] / sla_chart["Count"].sum() * 100).round(1)
            else:
                sla_chart["Percent"] = 0

            fig_sla = px.bar(
                sla_chart,
                x="SLA_Category",
                y="Count",
                text=sla_chart.apply(lambda row: f"{row['Count']} ({row['Percent']}%)", axis=1),
                title="SLA Distribution"
            )
            fig_sla.update_traces(textposition="outside", textfont=dict(size=14))
            st.plotly_chart(fig_sla, use_container_width=True, key="fig_sla")

            total_sla_count = int(sla_chart["Count"].sum()) if not sla_chart.empty else 0
            period_months = filtered_df["MonthYear"].nunique() if "MonthYear" in filtered_df.columns else 0
            st.write(f"**Total Count across all SLA slabs: {total_sla_count} ({period_months} months)**")
        period_months = filtered_df["MonthYear"].nunique() if "MonthYear" in filtered_df.columns else 0
        
        # Month-wise Count and Value
        st.subheader(f"Month-wise Count and Value ({period_months} Months)")
        month_trend = filtered_df.groupby("MonthYear")[["ACCOUNT NO", "Plan Value"]].agg(
            {"ACCOUNT NO": "count", "Plan Value": "sum"}
        ).reset_index()
        month_trend.rename(columns={"ACCOUNT NO": "Count", "Plan Value": "Value"}, inplace=True)

        if not month_trend.empty:
            month_trend["MonthYear_dt"] = pd.to_datetime(month_trend["MonthYear"], format="%b-%Y", errors="coerce")
            month_trend = month_trend.sort_values("MonthYear_dt")
            month_trend["MonthYear"] = month_trend["MonthYear_dt"].dt.strftime("%b-%Y")
            month_trend["Value_Display"] = month_trend["Value"].apply(format_value)

            fig_month = px.bar(
                month_trend,
                x="MonthYear",
                y="Count",
                text="Count",
                title="Month-wise Count and Value"
            )

            fig_month.update_traces(
                textposition="inside",
                textfont=dict(color="white", size=10),
                insidetextanchor="start",
                selector=dict(type="bar")
            )

            fig_month.add_scatter(
                x=month_trend["MonthYear"],
                y=month_trend["Value"],
                mode="lines+markers+text",
                name="Value",
                text=month_trend["Value_Display"],
                textposition="top center",
                textfont=dict(size=12, color="black"),
                yaxis="y2"
            )

            fig_month.update_layout(
                xaxis_title="Month",
                yaxis=dict(title="Count"),
                yaxis2=dict(title="Value", overlaying="y", side="right", showgrid=False),
                xaxis=dict(tickangle=45)
            )

            st.plotly_chart(fig_month, use_container_width=True, key="fig_month")

        total_count_month = month_trend["Count"].sum() if not month_trend.empty else 0
        period_months_month = month_trend["MonthYear"].nunique() if not month_trend.empty else 0
        st.write(f"**Total Count across all months: {total_count_month} ({period_months_month} months)**")
        
        if contribution_mode == "Installations + Winback":
            st.subheader(f"Month-wise Consolidated Installation + Winback Summary ({period_months_month} Months)")

            install_monthly = (
                filtered_df[filtered_df["Record_Type"] == "Installation"]
                .groupby("MonthYear")
                .agg(
                    Installation_Count=("ACCOUNT NO", "count"),
                    Installation_Value=("Plan Value", "sum")
                )
                .reset_index()
            )

            winback_monthly = (
                filtered_df[filtered_df["Record_Type"] == "Winback"]
                .groupby("MonthYear")
                .agg(
                    Winback_Count=("ACCOUNT NO", "count"),
                    Winback_Value=("Plan Value", "sum")
                )
                .reset_index()
            )

            consolidated = pd.merge(
                install_monthly,
                winback_monthly,
                on="MonthYear",
                how="outer"
            ).fillna(0)

            consolidated["MonthYear_dt"] = pd.to_datetime(consolidated["MonthYear"], format="%b-%Y", errors="coerce")
            consolidated = consolidated.sort_values("MonthYear_dt")

            consolidated["Total_Count"] = consolidated["Installation_Count"] + consolidated["Winback_Count"]
            consolidated["Total_Value"] = consolidated["Installation_Value"] + consolidated["Winback_Value"]

            # round values
            for col in ["Installation_Value", "Winback_Value", "Total_Value"]:
                consolidated[col] = consolidated[col].round(0)

            # total row
            total_row = pd.DataFrame({
                "MonthYear": ["Total"],
                "MonthYear_dt": [pd.NaT],
                "Installation_Count": [consolidated["Installation_Count"].sum()],
                "Installation_Value": [consolidated["Installation_Value"].sum()],
                "Winback_Count": [consolidated["Winback_Count"].sum()],
                "Winback_Value": [consolidated["Winback_Value"].sum()],
                "Total_Count": [consolidated["Total_Count"].sum()],
                "Total_Value": [consolidated["Total_Value"].sum()]
            })

            consolidated_display = pd.concat([consolidated, total_row], ignore_index=True)

            st.dataframe(
                consolidated_display[
                    [
                        "MonthYear",
                        "Installation_Count", "Installation_Value",
                        "Winback_Count", "Winback_Value",
                        "Total_Count", "Total_Value"
                    ]
                ],
                use_container_width=True
            )
        
            # City-level Executive Performance Table
        if city != "All" and execu == "All":
            st.subheader(f"City-level Executive Performance: {city}")

            city_df = filtered_df_base[filtered_df_base["City"] == city].copy()

            exec_status_table = st.selectbox(
                "Executive Status Filter for Table",
                ["All"] + sorted(city_df["EXEC_STATUS_FINAL"].dropna().astype(str).unique()) if "EXEC_STATUS_FINAL" in city_df.columns else ["All"],
                key="city_exec_status_table"
            )

            if exec_status_table != "All" and "EXEC_STATUS_FINAL" in city_df.columns:
                city_df = city_df[city_df["EXEC_STATUS_FINAL"] == exec_status_table]

            city_df["Exec_Display"] = city_df["EXEC_NAME_FINAL"].astype(str) + " (" + city_df["EXEC_STATUS_FINAL"].astype(str) + ")"

            city_exec = city_df.groupby(["Exec_Display", "MonthYear"]).agg(
                Count=("ACCOUNT NO", "count"),
                Value=("Plan Value", "sum")
            ).reset_index()

            city_exec["MonthYear_dt"] = pd.to_datetime(city_exec["MonthYear"], format="%b-%Y", errors="coerce")
            city_exec = city_exec.dropna(subset=["MonthYear_dt"]).sort_values("MonthYear_dt")
            city_exec["MonthYear"] = city_exec["MonthYear_dt"].dt.strftime("%b-%y")

            pivot_table = city_exec.pivot_table(
                index="Exec_Display",
                columns="MonthYear",
                values=["Count", "Value"],
                fill_value=0
            )

            months_sorted = sorted(city_exec["MonthYear"].unique(), key=lambda x: pd.to_datetime(x, format="%b-%y"))
            new_cols = []
            for m in months_sorted:
                new_cols.append(("Count", m))
                new_cols.append(("Value", m))
            pivot_table = pivot_table.reindex(columns=new_cols, fill_value=0)

            pivot_table.columns = [f"{col[1]} {col[0]}" for col in pivot_table.columns]

            totals = city_exec.groupby("Exec_Display")[["Count", "Value"]].sum()
            avgs = city_exec.groupby("Exec_Display")[["Count", "Value"]].mean().round(0)

            pivot_table["Total Count"] = totals["Count"]
            pivot_table["Total Value"] = totals["Value"]
            pivot_table["Avg/Month Count"] = avgs["Count"]
            pivot_table["Avg/Month Value"] = avgs["Value"]

            st.dataframe(pivot_table.sort_values("Total Count", ascending=False), use_container_width=True)

            csv_data = pivot_table.to_csv().encode("utf-8")
            st.download_button(
                "📥 Download Executive Table (CSV)",
                csv_data,
                file_name=f"Executive_Monthly_{city}.csv",
                mime="text/csv"
            )
            
            
            # City-level Ageing Summary

            exec_month_summary = generate_summary(
                city_df.rename(columns={"EXEC_NAME_FINAL": "Exec Name"}),
                level="Executive",
                group_col="Exec Name"
            )
            st.write(exec_month_summary)

        # Top Nodes
        period_months = filtered_df["MonthYear"].nunique() if "MonthYear" in filtered_df.columns else 0
        st.subheader(f"Top 10 Nodes by Count ({period_months} Months)")

        node_col = "INSTALLATION NODE"
        if node_col in filtered_df.columns:
            node_data = filtered_df.groupby([node_col, "City"])["ACCOUNT NO"].count().reset_index()
            node_data.rename(columns={"ACCOUNT NO": f"Count ({period_months} Months)"}, inplace=True)

            top_nodes = node_data.sort_values(f"Count ({period_months} Months)", ascending=False).head(10)

            fig_nodes = px.bar(
                top_nodes,
                x=node_col,
                y=f"Count ({period_months} Months)",
                color="City",
                text=top_nodes.apply(lambda row: f"{row['City']} - {row[f'Count ({period_months} Months)']}", axis=1),
                title=f"Top 10 Nodes by Count ({period_months} Months)"
            )
            fig_nodes.update_traces(textposition="outside", textfont=dict(size=12))
            fig_nodes.update_layout(xaxis_title="Node", yaxis_title="Count", xaxis_tickangle=45)
            st.plotly_chart(fig_nodes, use_container_width=True, key="fig_nodes")

            st.subheader(f"Top 10 Nodes Table ({period_months} Months)")
            st.dataframe(top_nodes.reset_index(drop=True).rename_axis(None), use_container_width=True)
        else:
            st.warning("⚠️ Column 'INSTALLATION NODE' not found in dataset.")

        # Executive Ageing & Performance Expectation - moved to bottom
        if "EXEC_NAME_FINAL" in filtered_df_base.columns and "Exec_Ageing_Bucket" in filtered_df_base.columns:
            st.subheader(f"Executive Ageing Wise Performance({period_months} Months)")
            perf_eval = (
                filtered_df_base.groupby(
                    ["EXEC_NAME_FINAL", "EXEC_STATUS_FINAL", "DEPARTMENT_FINAL", "Exec_Ageing_Bucket", "Target_Avg_Month"]
                )
                .agg(
                    Total=("ACCOUNT NO", "count"),
                    Installation=("Record_Type", lambda x: (x == "Installation").sum()),
                    Winback=("Record_Type", lambda x: (x == "Winback").sum()),
                    Months=("MonthYear", "nunique"),
                    Value=("Plan Value", "sum")
                )
                .reset_index()
            )

            perf_eval["Avg/Month Count"] = (perf_eval["Total"] / perf_eval["Months"]).round(1)
            perf_eval["Avg/Month Value"] = (perf_eval["Value"] / perf_eval["Months"]).round(0)
            perf_eval["Expectation Status"] = perf_eval.apply(
                lambda row: expectation_label(row["Avg/Month Count"], row["Target_Avg_Month"]),
                axis=1
            )

            perf_eval = perf_eval.rename(columns={
                "EXEC_NAME_FINAL": "Exec Name",
                "EXEC_STATUS_FINAL": "Status",
                "DEPARTMENT_FINAL": "Department",
                "Exec_Ageing_Bucket": "Ageing Bucket",
                "Target_Avg_Month": "Target Avg/Month"
            })

            ageing_sort_map = {
                "Less Than 3 Months": 1,
                "4-12 Months": 2,
                "Above 12 Months": 3,
                "Missing": 4
            }
            perf_eval["Ageing_Sort"] = perf_eval["Ageing Bucket"].map(ageing_sort_map)
            perf_eval = perf_eval.sort_values(["Ageing_Sort", "Avg/Month Count"], ascending=[True, False]).drop(columns=["Ageing_Sort"])

            st.dataframe(perf_eval, use_container_width=True)

        # Monthly Executive Contribution - moved to bottom
            st.subheader(f"Executive Ageing Summary ({period_months} Months)")

            ageing_eval = (
            filtered_df_base.groupby(
                ["EXEC_NAME_FINAL", "EXEC_STATUS_FINAL", "Exec_Ageing_Bucket", "Target_Avg_Month"]
            )
            .agg(
                Total=("ACCOUNT NO", "count"),
                Months=("MonthYear", "nunique"),
                Value=("Plan Value", "sum")
            )
            .reset_index()
        )

        ageing_eval["Avg/Month Count"] = (ageing_eval["Total"] / ageing_eval["Months"]).round(1)
        ageing_eval["Avg/Month Value"] = (ageing_eval["Value"] / ageing_eval["Months"]).round(0)
        ageing_eval["Expectation Status"] = ageing_eval.apply(
            lambda row: expectation_label(row["Avg/Month Count"], row["Target_Avg_Month"]),
            axis=1
        )

        ageing_summary_table = (
            ageing_eval.groupby(["Exec_Ageing_Bucket", "Expectation Status"])
            .agg(Employee_Count=("EXEC_NAME_FINAL", "nunique"))
            .reset_index()
        )

        total_emp_by_bucket = (
            ageing_eval.groupby("Exec_Ageing_Bucket")
            .agg(Total_Employees=("EXEC_NAME_FINAL", "nunique"))
            .reset_index()
        )

        ageing_summary_table = ageing_summary_table.merge(total_emp_by_bucket, on="Exec_Ageing_Bucket", how="left")
        ageing_summary_table["Percent Contribution"] = (
            ageing_summary_table["Employee_Count"] / ageing_summary_table["Total_Employees"] * 100
        ).round(1)

        ageing_summary_pivot = ageing_summary_table.pivot_table(
            index="Exec_Ageing_Bucket",
            columns="Expectation Status",
            values="Employee_Count",
            fill_value=0
        ).reset_index()

        for col in ["Below Expectations", "Meeting Expectations", "Exceptional", "Missing"]:
            if col not in ageing_summary_pivot.columns:
                ageing_summary_pivot[col] = 0

        ageing_summary_pivot = ageing_summary_pivot.merge(total_emp_by_bucket, on="Exec_Ageing_Bucket", how="left")

        ageing_pct = ageing_summary_table.pivot_table(
            index="Exec_Ageing_Bucket",
            columns="Expectation Status",
            values="Percent Contribution",
            fill_value=0
        ).reset_index()

        for col in ["Below Expectations", "Meeting Expectations", "Exceptional", "Missing"]:
            if col not in ageing_pct.columns:
                ageing_pct[col] = 0

        ageing_pct = ageing_pct.rename(columns={
            "Below Expectations": "Below %",
            "Meeting Expectations": "Meeting %",
            "Exceptional": "Exceptional %",
            "Missing": "Missing %"
        })

        ageing_summary_pivot = ageing_summary_pivot.merge(ageing_pct, on="Exec_Ageing_Bucket", how="left")

        ageing_order = ["Less Than 3 Months", "4-12 Months", "Above 12 Months", "Missing"]
        ageing_summary_pivot["Exec_Ageing_Bucket"] = pd.Categorical(
            ageing_summary_pivot["Exec_Ageing_Bucket"],
            categories=ageing_order,
            ordered=True
        )
        ageing_summary_pivot = ageing_summary_pivot.sort_values("Exec_Ageing_Bucket")
         
        total_summary = pd.DataFrame({
            "Exec_Ageing_Bucket": ["Total"],
            "Below Expectations": [ageing_summary_pivot["Below Expectations"].sum() if "Below Expectations" in ageing_summary_pivot.columns else 0],
            "Meeting Expectations": [ageing_summary_pivot["Meeting Expectations"].sum() if "Meeting Expectations" in ageing_summary_pivot.columns else 0],
            "Exceptional": [ageing_summary_pivot["Exceptional"].sum() if "Exceptional" in ageing_summary_pivot.columns else 0],
            "Missing": [ageing_summary_pivot["Missing"].sum() if "Missing" in ageing_summary_pivot.columns else 0],
            "Total_Employees": [ageing_summary_pivot["Total_Employees"].sum() if "Total_Employees" in ageing_summary_pivot.columns else 0],
            "Below %": [round(ageing_summary_pivot["Below Expectations"].sum() / ageing_summary_pivot["Total_Employees"].sum() * 100, 1) if "Below Expectations" in ageing_summary_pivot.columns and ageing_summary_pivot["Total_Employees"].sum() > 0 else 0],
            "Meeting %": [round(ageing_summary_pivot["Meeting Expectations"].sum() / ageing_summary_pivot["Total_Employees"].sum() * 100, 1) if "Meeting Expectations" in ageing_summary_pivot.columns and ageing_summary_pivot["Total_Employees"].sum() > 0 else 0],
            "Exceptional %": [round(ageing_summary_pivot["Exceptional"].sum() / ageing_summary_pivot["Total_Employees"].sum() * 100, 1) if "Exceptional" in ageing_summary_pivot.columns and ageing_summary_pivot["Total_Employees"].sum() > 0 else 0],
            "Missing %": [round(ageing_summary_pivot["Missing"].sum() / ageing_summary_pivot["Total_Employees"].sum() * 100, 1) if "Missing" in ageing_summary_pivot.columns and ageing_summary_pivot["Total_Employees"].sum() > 0 else 0]
        })
        ageing_summary_pivot = pd.concat([ageing_summary_pivot, total_summary], ignore_index=True)    

        st.dataframe(ageing_summary_pivot, use_container_width=True)

    # ==================================================
    # TAB 2: Customer Profiling
    # ==================================================
    with tab2:
        if "VALIDITY In Months" in filtered_df.columns:
            validity_pattern = filtered_df.groupby(["VALIDITY In Months"])["ACCOUNT NO"].count().reset_index()
            validity_pattern.rename(columns={"ACCOUNT NO": "Count"}, inplace=True)
            total_validity = validity_pattern["Count"].sum()
            validity_pattern["Percent"] = (validity_pattern["Count"] / total_validity * 100).round(1) if total_validity > 0 else 0
            validity_pattern["Validity Label"] = validity_pattern["VALIDITY In Months"].astype(str) + " Months"
            validity_pattern = validity_pattern[validity_pattern["Percent"] >= 1]

            fig_validity = px.pie(
                validity_pattern,
                names="Validity Label",
                values="Count",
                title="Validity-wise Distribution"
            )
            fig_validity.update_traces(textinfo="label+percent", textfont=dict(size=14))
            st.plotly_chart(fig_validity, use_container_width=True, key="fig_validity")

        if "ARPU_BUCKET" in filtered_df.columns:
            slab_order = ["upto 300", "301-500", "501-750", "751+", "Missing"]
            arpu_counts = filtered_df.groupby("ARPU_BUCKET")["ACCOUNT NO"].count().reset_index()
            arpu_counts.rename(columns={"ACCOUNT NO": "Count"}, inplace=True)

            arpu_counts["ARPU_BUCKET"] = pd.Categorical(
                arpu_counts["ARPU_BUCKET"],
                categories=slab_order,
                ordered=True
            )
            arpu_counts = arpu_counts.sort_values("ARPU_BUCKET")

            total_count_arpu = arpu_counts["Count"].sum()
            arpu_counts["Percent"] = (arpu_counts["Count"] / total_count_arpu * 100).round(0).astype(int) if total_count_arpu > 0 else 0
            arpu_counts["Label"] = arpu_counts.apply(lambda row: f"{row['Count']} ({row['Percent']}%)", axis=1)

            fig_arpu = px.bar(
                arpu_counts,
                x="ARPU_BUCKET",
                y="Count",
                color="ARPU_BUCKET",
                title="ARPU Slabs Distribution",
                text="Label",
                category_orders={"ARPU_BUCKET": slab_order}
            )
            fig_arpu.update_traces(textposition="outside")
            st.plotly_chart(fig_arpu, use_container_width=True, key="fig_arpu")
            st.write(f"**Total Count across ARPU slabs: {int(total_count_arpu)}**")

        if "SPEED (Mbps)" in filtered_df.columns:
            speed_pattern = filtered_df.groupby(["SPEED (Mbps)"])["ACCOUNT NO"].count().reset_index()
            speed_pattern.rename(columns={"ACCOUNT NO": "Count"}, inplace=True)
            total_speed = speed_pattern["Count"].sum()
            speed_pattern["Percent"] = (speed_pattern["Count"] / total_speed * 100).round(1) if total_speed > 0 else 0
            speed_pattern["Speed Label"] = speed_pattern["SPEED (Mbps)"].astype(str) + " Mbps"
            speed_pattern = speed_pattern[speed_pattern["Percent"] >= 1]

            fig_speed = px.pie(
                speed_pattern,
                names="Speed Label",
                values="Count",
                title="Speed-wise Distribution"
            )
            fig_speed.update_traces(textinfo="label+percent", textfont=dict(size=14))
            st.plotly_chart(fig_speed, use_container_width=True, key="fig_speed")

        if "NETWORK TYPE" in filtered_df.columns:
            network_counts = filtered_df.groupby("NETWORK TYPE")["ACCOUNT NO"].count().reset_index().rename(columns={"ACCOUNT NO": "Count"})
            total_count_network = network_counts["Count"].sum()
            network_counts["Percent"] = (network_counts["Count"] / total_count_network * 100).round(0).astype(int) if total_count_network > 0 else 0
            network_counts = network_counts[network_counts["Percent"] >= 1]

            fig_network = px.pie(
                network_counts,
                names="NETWORK TYPE",
                values="Count",
                title="Network Type Contribution (≥1%)",
                hover_data=["Percent"]
            )
            fig_network.update_traces(textinfo="label+percent")
            st.plotly_chart(fig_network, use_container_width=True, key="fig_network")

        # Restored customer status bar chart
        if "CUSTOMER CURRENT STATUS" in filtered_df.columns and "MonthYear" in filtered_df.columns:
            filtered_df_status = filtered_df.copy()
            filtered_df_status["MonthYear_dt"] = pd.to_datetime(
                filtered_df_status["MonthYear"], format="%b-%Y", errors="coerce"
            )
            filtered_df_status = filtered_df_status.dropna(subset=["MonthYear_dt"])

            if not filtered_df_status.empty:
                last_12 = filtered_df_status["MonthYear_dt"].max() - pd.DateOffset(months=12)
                df_last12 = filtered_df_status[filtered_df_status["MonthYear_dt"] >= last_12]

                if not df_last12.empty:
                    status_counts = (
                        df_last12.groupby(["MonthYear_dt", "CUSTOMER CURRENT STATUS"])["ACCOUNT NO"]
                        .count()
                        .reset_index()
                        .rename(columns={"ACCOUNT NO": "Count"})
                    )

                    status_counts = status_counts.sort_values("MonthYear_dt")
                    status_counts["MonthYear_str"] = status_counts["MonthYear_dt"].dt.strftime("%b-%Y")

                    totals = (
                        status_counts.groupby("MonthYear_str")["Count"]
                        .sum()
                        .reset_index()
                        .rename(columns={"Count": "Total"})
                    )

                    status_counts = status_counts.merge(totals, on="MonthYear_str", how="left")
                    status_counts["Percent"] = (status_counts["Count"] / status_counts["Total"] * 100).round(1)
                    status_counts["Label"] = status_counts.apply(lambda row: f"{row['Count']} ({row['Percent']}%)", axis=1)

                    fig_status = px.bar(
                        status_counts,
                        x="MonthYear_str",
                        y="Count",
                        color="CUSTOMER CURRENT STATUS",
                        title="Customer Current Status - Last 12 Months",
                        text="Label"
                    )
                    fig_status.update_traces(textposition="inside")

                    for _, row in totals.iterrows():
                        fig_status.add_annotation(
                            x=row["MonthYear_str"],
                            y=row["Total"],
                            text=str(row["Total"]),
                            showarrow=False,
                            font=dict(color="black", size=12, family="Arial"),
                            yshift=10
                        )

                    st.plotly_chart(fig_status, use_container_width=True, key="fig_status")
                else:
                    st.info("No customer status data available for the selected filters in the last 12 months.")
    # ==================================================
    # TAB 3: CEO & Sales Head
    # ==================================================
    with tab3:
        st.header("CEO / National Sales Manager Analytics")

        period_months = source_df["MonthYear"].nunique() if "MonthYear" in source_df.columns else 0

        total_count = filtered_df["ACCOUNT NO"].nunique() if "ACCOUNT NO" in filtered_df.columns else 0
        total_value = filtered_df["Plan Value"].sum() if "Plan Value" in filtered_df.columns else 0
        avg_count = filtered_df.groupby("MonthYear")["ACCOUNT NO"].nunique().mean() if "MonthYear" in filtered_df.columns and not filtered_df.empty else 0
        avg_value = filtered_df.groupby("MonthYear")["Plan Value"].sum().mean() if "MonthYear" in filtered_df.columns and not filtered_df.empty else 0

        avg_count = 0 if pd.isna(avg_count) else round(avg_count, 0)
        avg_value = 0 if pd.isna(avg_value) else round(avg_value, 0)

        kpi_cols = st.columns(4)
        kpi_cols[0].metric("Total Count", format_value(total_count))
        kpi_cols[1].metric("Total Value", f"₹{format_value(total_value)}")
        kpi_cols[2].metric("Avg Count/Month", format_value(avg_count))
        kpi_cols[3].metric("Avg Value/Month", f"₹{format_value(avg_value)}")

        if "EXEC_STATUS_FINAL" in source_df.columns:
            ceo_summary = (
                source_df.groupby("EXEC_STATUS_FINAL")
                .agg(
                    Count=("ACCOUNT NO", "count"),
                    Value=("Plan Value", "sum")
                )
                .reset_index()
            )

            exec_counts = source_df.groupby("EXEC_STATUS_FINAL")["SALES CODE"].nunique().reset_index()
            exec_counts.rename(columns={"SALES CODE": "Exec_Count"}, inplace=True)

            ceo_summary = ceo_summary.merge(exec_counts, on="EXEC_STATUS_FINAL", how="left")
            ceo_summary["Avg Count/Exec"] = (ceo_summary["Count"] / ceo_summary["Exec_Count"]).round(0)
            ceo_summary["Avg Value/Exec"] = (ceo_summary["Value"] / ceo_summary["Exec_Count"]).round(0)

            st.subheader("Executive Status KPIs")

            ceo_summary_display = ceo_summary.copy()
            ceo_summary_display["Total Value"] = ceo_summary_display["Value"].apply(format_currency_compact)
            ceo_summary_display["Avg Value/Exec"] = ceo_summary_display["Avg Value/Exec"].apply(format_currency_compact)

            st.dataframe(
                ceo_summary_display[
                    ["EXEC_STATUS_FINAL", "Count", "Total Value", "Exec_Count", "Avg Count/Exec", "Avg Value/Exec"]
                ].rename(columns={
                    "EXEC_STATUS_FINAL": "Status",
                    "Count": "Total Count",
                    "Exec_Count": "No. of Executives"
                }),
                use_container_width=True
            )

        with st.expander("Top Executives by Count / Value"):
            top_exec_count = (
                source_df.groupby(["City", "EXEC_NAME_FINAL"], as_index=False)
                .agg({"ACCOUNT NO": "count"})
            )
            top_exec_count.rename(columns={"ACCOUNT NO": f"Total Count ({period_months} Months)"}, inplace=True)
            top_exec_count = top_exec_count.sort_values(f"Total Count ({period_months} Months)", ascending=False).head(20)
            st.subheader(f"Top 20 Executives by Count ({period_months} Months)")
            st.dataframe(top_exec_count.reset_index(drop=True).rename_axis(None), use_container_width=True)

            top_exec_value = (
                source_df.groupby(["City", "EXEC_NAME_FINAL"], as_index=False)
                .agg({"Plan Value": "sum"})
            )
            top_exec_value["Plan Value"] = pd.to_numeric(top_exec_value["Plan Value"], errors="coerce").fillna(0)
            top_exec_value["Plan Value"] = (top_exec_value["Plan Value"] / 100000).round(2)
            top_exec_value.rename(columns={"Plan Value": f"Total Value (₹ Lacs, {period_months} Months)"}, inplace=True)
            top_exec_value = top_exec_value.sort_values(f"Total Value (₹ Lacs, {period_months} Months)", ascending=False).head(20)

            st.subheader(f"Top 20 Executives by Value (₹ Lacs, {period_months} Months)")
            st.dataframe(top_exec_value.reset_index(drop=True).rename_axis(None), use_container_width=True)

        with st.expander("Top Cities by Sales / Collections"):
            top_cities_sales = source_df.groupby("City")["Plan Value"].sum().reset_index()
            top_cities_sales["Plan Value"] = (top_cities_sales["Plan Value"] / 1_000_000).round(2)
            top_cities_sales.rename(columns={"Plan Value": f"Total Sales Value (₹ Millions, {period_months} Months)"}, inplace=True)
            top_cities_sales = top_cities_sales.sort_values(f"Total Sales Value (₹ Millions, {period_months} Months)", ascending=False).head(5)
            st.subheader(f"Top 5 Cities by Total Sales Value (₹ Millions, {period_months} Months)")
            st.dataframe(top_cities_sales.reset_index(drop=True).rename_axis(None), use_container_width=True)

            fig_cities_sales = px.pie(
                top_cities_sales,
                names="City",
                values=f"Total Sales Value (₹ Millions, {period_months} Months)",
                title="Top Cities by Total Sales Value"
            )
            st.plotly_chart(fig_cities_sales, use_container_width=True, key="fig_cities_sales")

            X_months = 3
            cutoff_date = pd.to_datetime("today") - pd.DateOffset(months=X_months)
            recent_df = source_df[source_df["INSTALLATION DATE"] >= cutoff_date]
            top_cities_recent = recent_df.groupby("City")["Plan Value"].sum().reset_index()
            top_cities_recent["Plan Value"] = (top_cities_recent["Plan Value"] / 1_000_000).round(2)
            top_cities_recent.rename(columns={"Plan Value": f"Total Sales Value (₹ Millions, Last {X_months} Months)"}, inplace=True)
            top_cities_recent = top_cities_recent.sort_values(f"Total Sales Value (₹ Millions, Last {X_months} Months)", ascending=False).head(5)
            st.subheader(f"Top 5 Cities by Collections in Last {X_months} Months (₹ Millions)")
            st.dataframe(top_cities_recent.reset_index(drop=True).rename_axis(None), use_container_width=True)

            fig_cities_recent = px.bar(
                top_cities_recent,
                x="City",
                y=f"Total Sales Value (₹ Millions, Last {X_months} Months)",
                text=f"Total Sales Value (₹ Millions, Last {X_months} Months)",
                title=f"Top Cities by Collections in Last {X_months} Months"
            )
            fig_cities_recent.update_traces(textposition="outside")
            st.plotly_chart(fig_cities_recent, use_container_width=True, key="fig_cities_recent")

        with st.expander("Top Cities Selling High-Speed Plans (≥100 Mbps)"):
            high_speed_df = source_df[source_df["SPEED (Mbps)"] >= 100] if "SPEED (Mbps)" in source_df.columns else pd.DataFrame()
            top_cities_speed = high_speed_df.groupby("City")["ACCOUNT NO"].count().reset_index()
            top_cities_speed.rename(columns={"ACCOUNT NO": f"High-Speed Count ({period_months} Months)"}, inplace=True)
            top_cities_speed = top_cities_speed.sort_values(f"High-Speed Count ({period_months} Months)", ascending=False).head(5)
            st.subheader(f"Top 5 Cities Selling 100 Mbps+ Plans ({period_months} Months)")
            st.dataframe(top_cities_speed.reset_index(drop=True).rename_axis(None), use_container_width=True)

            fig_cities_speed = px.bar(
                top_cities_speed,
                x="City",
                y=f"High-Speed Count ({period_months} Months)",
                text=f"High-Speed Count ({period_months} Months)",
                title="Top Cities Selling 100 Mbps+ Plans"
            )
            fig_cities_speed.update_traces(textposition="outside")
            st.plotly_chart(fig_cities_speed, use_container_width=True, key="fig_cities_speed")

        with st.expander("Executives with Low Count (<3 per Month in Last X Months)"):
            X_months = 3
            cutoff_date = pd.to_datetime("today") - pd.DateOffset(months=X_months)
            recent_exec = source_df[source_df["INSTALLATION DATE"] >= cutoff_date]

            exec_monthly = recent_exec.groupby(["City", "EXEC_NAME_FINAL", "MonthYear"])["ACCOUNT NO"].count().reset_index()
            low_execs = exec_monthly.groupby(["City", "EXEC_NAME_FINAL"])["ACCOUNT NO"].mean().reset_index()
            low_execs = low_execs[low_execs["ACCOUNT NO"] < 3]
            low_execs.rename(columns={"ACCOUNT NO": f"Avg Count/Month (Last {X_months} Months)"}, inplace=True)

            st.subheader(f"Executives Averaging <3 Count/Month in Last {X_months} Months")
            st.dataframe(low_execs.reset_index(drop=True).rename_axis(None), use_container_width=True)

    # --------------------------------------------------
    # Footer info
    # --------------------------------------------------
    latest_month = (
        source_df["MonthYear"].dropna().iloc[-1]
        if "MonthYear" in source_df.columns and not source_df["MonthYear"].dropna().empty
        else "N/A"
    )
    st.caption(f"Latest month available in data: {latest_month}")
