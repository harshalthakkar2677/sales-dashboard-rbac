import os
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import calendar
import tempfile
from email.mime.base import MIMEBase
from email import encoders

# --------------------------------------------------
# File Paths
# --------------------------------------------------
DATA_FILE = "New Registration Report.csv"
WINBACK_FILE = "New Winback Report.csv"
ACCESS_FILE = "Access Master.xlsx"

# --------------------------------------------------
# SMTP Configuration - UPDATE THESE
# --------------------------------------------------
SMTP_HOST = "officemail.youbroadband.in"
SMTP_PORT = 465
SMTP_LOGIN_USER = "hthakkar"
SENDER_EMAIL = "harshal.thakkar@youbroadband.co.in"
SENDER_PASSWORD = "A_b4e2756179"

# --------------------------------------------------
# Helper Functions
# --------------------------------------------------
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
        "sales direct": "Sales Direct",
        "esg - sales": "ESG - Sales",
        "sales": "Sales",
        "online": "Online",
        "retention": "Retention",
        "technical": "Technical"
    }
    return mapping.get(x, str(x).strip().title())

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
    if pd.isna(avg_sales) or pd.isna(target) or target == 0:
        return "Missing"

    lower_band = target * 0.90
    upper_band = target * 1.10

    if avg_sales < lower_band:
        return "Below Expectations"
    elif avg_sales <= upper_band:
        return "Meeting Expectations"
    else:
        return "Exceptional"

def parse_date_series(series):
    s = series.astype(str).str.strip()
    s = s.replace(["", "nan", "None", "NA", "N/A", "Online"], np.nan)

    p1 = pd.to_datetime(s, format="%d-%m-%Y", errors="coerce")
    p2 = pd.to_datetime(s, format="%d/%m/%Y", errors="coerce")
    p3 = pd.to_datetime(s, format="%Y-%m-%d", errors="coerce")
    p4 = pd.to_datetime(s, format="%d-%m-%y %H:%M", errors="coerce")
    p5 = pd.to_datetime(s, format="%d-%b-%y", errors="coerce")
    p6 = pd.to_datetime(s, format="%d-%b-%Y", errors="coerce")

    parsed = p1.fillna(p2).fillna(p3).fillna(p4).fillna(p5).fillna(p6)

    mask = parsed.isna() & s.notna()
    if mask.any():
        parsed.loc[mask] = pd.to_datetime(s.loc[mask], dayfirst=True, errors="coerce")

    return parsed

def format_currency(num):
    if pd.isna(num):
        return "₹0"
    return f"₹{num:,.0f}"

def safe_divide(a, b):
    if b in [0, None] or pd.isna(b):
        return 0
    return a / b

# --------------------------------------------------
# Load Access Master
# --------------------------------------------------
def load_access_master():
    access_df = pd.read_excel(ACCESS_FILE)
    access_df.columns = access_df.columns.str.strip()
    access_df["UserID"] = access_df["UserID"].astype(str).str.strip()
    access_df["Username"] = access_df["Username"].astype(str).str.strip()
    access_df["Role"] = access_df["Role"].astype(str).str.strip()
    access_df["Region/City"] = access_df["Region/City"].astype(str).str.strip()
    access_df["Email"] = access_df["Email"].astype(str).str.strip()
    return access_df

# --------------------------------------------------
# Load and Prepare Installation Data
# --------------------------------------------------
def load_installation_data():
    df = pd.read_csv(DATA_FILE)
    df.columns = df.columns.str.strip()

    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).str.strip()

    if "CREATION DATE" in df.columns:
        df["CREATION DATE"] = parse_date_series(df["CREATION DATE"])

    if "INSTALLATION DATE" in df.columns:
        df["INSTALLATION DATE"] = parse_date_series(df["INSTALLATION DATE"])
        df["MonthStart"] = df["INSTALLATION DATE"].dt.to_period("M").dt.to_timestamp()
        df["MonthYear"] = df["MonthStart"].dt.strftime("%b-%Y")
        df["Year"] = df["INSTALLATION DATE"].dt.year
    else:
        df["MonthStart"] = pd.NaT
        df["MonthYear"] = pd.NA
        df["Year"] = pd.NA

    if "SALES EXEC DOJ" in df.columns:
        df["SALES EXEC DOJ"] = parse_date_series(df["SALES EXEC DOJ"])
    else:
        df["SALES EXEC DOJ"] = pd.NaT

    if "Plan Value" in df.columns:
        df["Plan Value"] = pd.to_numeric(df["Plan Value"], errors="coerce").fillna(0)
    else:
        df["Plan Value"] = 0

    if "ARPU" in df.columns:
        df["ARPU"] = pd.to_numeric(df["ARPU"], errors="coerce")
        df["ARPU_BUCKET"] = df["ARPU"].apply(arpu_bucket)
    else:
        df["ARPU"] = np.nan
        df["ARPU_BUCKET"] = "Missing"

    if "SALES EXEC STATUS" in df.columns:
        df["EXEC_STATUS_CLEAN"] = df["SALES EXEC STATUS"].apply(normalize_exec_status)
    else:
        df["EXEC_STATUS_CLEAN"] = "Missing"

    if "DEPARTMENT" in df.columns:
        df["DEPARTMENT_CLEAN"] = df["DEPARTMENT"].apply(normalize_department)
    else:
        df["DEPARTMENT_CLEAN"] = "Missing"

    if "Employee City" in df.columns:
        df["Employee City"] = df["Employee City"].astype(str).str.strip()
        df["Employee City"] = df["Employee City"].replace(["", "nan", "None"], pd.NA)
    else:
        df["Employee City"] = pd.NA

    if "SALES CODE" in df.columns:
        df["SALES CODE"] = df["SALES CODE"].astype(str).str.strip()

    if "SALES EXEC NAME" in df.columns:
        df["SALES EXEC NAME"] = df["SALES EXEC NAME"].astype(str).str.strip()

    df["Record_Type"] = "Installation"
    return df

# --------------------------------------------------
# Load and Prepare Winback Data
# --------------------------------------------------
def load_winback_data():
    if not os.path.exists(WINBACK_FILE):
        return pd.DataFrame()

    wb = pd.read_csv(WINBACK_FILE)
    wb.columns = wb.columns.str.strip()

    for col in wb.columns:
        if wb[col].dtype == "object":
            wb[col] = wb[col].astype(str).str.strip()

    if "INSTALLATION DATE" in wb.columns:
        wb["INSTALLATION DATE"] = parse_date_series(wb["INSTALLATION DATE"])
        wb["MonthStart"] = wb["INSTALLATION DATE"].dt.to_period("M").dt.to_timestamp()
        wb["MonthYear"] = wb["MonthStart"].dt.strftime("%b-%Y")
        wb["Year"] = wb["INSTALLATION DATE"].dt.year
    else:
        wb["MonthStart"] = pd.NaT
        wb["MonthYear"] = pd.NA
        wb["Year"] = pd.NA

    if "SALES EXEC DOJ" in wb.columns:
        wb["SALES EXEC DOJ"] = parse_date_series(wb["SALES EXEC DOJ"])
    else:
        wb["SALES EXEC DOJ"] = pd.NaT

    if "Plan Value" in wb.columns:
        wb["Plan Value"] = pd.to_numeric(wb["Plan Value"], errors="coerce").fillna(0)
    else:
        wb["Plan Value"] = 0

    if "ARPU" in wb.columns:
        wb["ARPU"] = pd.to_numeric(wb["ARPU"], errors="coerce")
        wb["ARPU_BUCKET"] = wb["ARPU"].apply(arpu_bucket)
    else:
        wb["ARPU"] = np.nan
        wb["ARPU_BUCKET"] = "Missing"

    if "SALES EXEC STATUS" in wb.columns:
        wb["EXEC_STATUS_CLEAN"] = wb["SALES EXEC STATUS"].apply(normalize_exec_status)
    else:
        wb["EXEC_STATUS_CLEAN"] = "Missing"

    if "DEPARTMENT" in wb.columns:
        wb["DEPARTMENT_CLEAN"] = wb["DEPARTMENT"].apply(normalize_department)
    else:
        wb["DEPARTMENT_CLEAN"] = "Missing"

    if "Employee City" in wb.columns:
        wb["Employee City"] = wb["Employee City"].astype(str).str.strip()
        wb["Employee City"] = wb["Employee City"].replace(["", "nan", "None"], pd.NA)
    else:
        wb["Employee City"] = pd.NA

    if "SALES CODE" in wb.columns:
        wb["SALES CODE"] = wb["SALES CODE"].astype(str).str.strip()

    if "SALES EXEC NAME" in wb.columns:
        wb["SALES EXEC NAME"] = wb["SALES EXEC NAME"].astype(str).str.strip()

    wb["Record_Type"] = "Winback"
    return wb

# --------------------------------------------------
# Build Employee Master from Installation Data
# --------------------------------------------------
def build_employee_master(df):
    sort_col = "INSTALLATION DATE" if "INSTALLATION DATE" in df.columns else "CREATION DATE"

    emp_master = (
        df.sort_values(sort_col)
          .dropna(subset=["SALES CODE"])
          .groupby("SALES CODE", as_index=False)
          .last()[["SALES CODE", "SALES EXEC NAME", "EXEC_STATUS_CLEAN", "DEPARTMENT_CLEAN", "SALES EXEC DOJ", "Employee City"]]
          .rename(columns={
              "SALES EXEC NAME": "EXEC_NAME_MASTER",
              "EXEC_STATUS_CLEAN": "EXEC_STATUS_MASTER",
              "DEPARTMENT_CLEAN": "DEPARTMENT_MASTER",
              "SALES EXEC DOJ": "DOJ_MASTER",
              "Employee City": "EXEC_CITY_MASTER"
          })
    )

    return emp_master

# --------------------------------------------------
# Apply Employee Mapping
# --------------------------------------------------
def apply_employee_mapping(base_df, emp_master):
    out = base_df.copy()

    out = out.drop(
        columns=[
            c for c in [
                "EXEC_NAME_MASTER",
                "EXEC_STATUS_MASTER",
                "DEPARTMENT_MASTER",
                "DOJ_MASTER",
                "EXEC_CITY_MASTER"
            ] if c in out.columns
        ],
        errors="ignore"
    )

    out = out.merge(emp_master, on="SALES CODE", how="left")

    out["EXEC_NAME_FINAL"] = out["EXEC_NAME_MASTER"].fillna(out["SALES EXEC NAME"])
    out["EXEC_STATUS_FINAL"] = out["EXEC_STATUS_MASTER"].fillna(out["EXEC_STATUS_CLEAN"])
    out["DEPARTMENT_FINAL"] = out["DEPARTMENT_MASTER"].fillna(out["DEPARTMENT_CLEAN"])
    out["SALES EXEC DOJ FINAL"] = out["DOJ_MASTER"]
    out["EXEC_CITY_FINAL"] = out["EXEC_CITY_MASTER"].fillna(out["Employee City"]).fillna("Unknown")

    # IMPORTANT FIX
    out["EXEC_KEY"] = out["SALES CODE"].astype(str).str.strip()

    return out


# --------------------------------------------------
# Add Ageing / Target / Expectation
# --------------------------------------------------
def add_expectation_fields(df):
    out = df.copy()
    today = pd.Timestamp.today().normalize()

    out["Exec_Ageing_Months"] = ((today - out["SALES EXEC DOJ FINAL"]).dt.days / 30.44).round(1)
    out["Exec_Ageing_Bucket"] = out["Exec_Ageing_Months"].apply(ageing_bucket)
    out["Target_Avg_Month"] = out["Exec_Ageing_Bucket"].apply(target_by_ageing)
    return out

# --------------------------------------------------
# Build Master Combined Dataset
# --------------------------------------------------
def build_master_dataset():
    df = load_installation_data()
    emp_master = build_employee_master(df)
    df = apply_employee_mapping(df, emp_master)
    df = add_expectation_fields(df)

    wb = load_winback_data()
    if not wb.empty:
        wb = apply_employee_mapping(wb, emp_master)
        wb = add_expectation_fields(wb)

    combined = pd.concat([df, wb], ignore_index=True) if not wb.empty else df.copy()
    return df, wb, combined, emp_master

# --------------------------------------------------
# Build Actionable Base
# --------------------------------------------------
def build_actionable_base(combined_df):
    action_df = combined_df.copy()
    perf_df = build_performance_base(combined_df)
    
    action_df = action_df[
        (action_df["DEPARTMENT_FINAL"] == "Sales") &
        (action_df["EXEC_STATUS_FINAL"] == "Active")
    ].copy()

    action_df["Installation"] = (action_df["Record_Type"] == "Installation").astype(int)
    action_df["Winback"] = (action_df["Record_Type"] == "Winback").astype(int)

    return action_df

def build_performance_base(combined_df):
    perf_df = combined_df.copy()

    # performance totals for MIS / manager view = Sales records in scope
    perf_df = perf_df[
        perf_df["DEPARTMENT_FINAL"] == "Sales"
    ].copy()

    perf_df["Installation"] = (perf_df["Record_Type"] == "Installation").astype(int)
    perf_df["Winback"] = (perf_df["Record_Type"] == "Winback").astype(int)

    return perf_df

# --------------------------------------------------
# Main Data Load
# --------------------------------------------------
access_df = load_access_master()
inst_df, winback_df, combined_df, emp_master = build_master_dataset()

perf_df = build_performance_base(combined_df)
action_df = build_actionable_base(combined_df)

print("\n=== PERFORMANCE DF VALIDATION ===")
print("Performance rows:", len(perf_df))
print("Performance columns:", perf_df.columns.tolist())

print("\n=== ACTION DF VALIDATION ===")
print("Columns in action_df:")
print(action_df.columns.tolist())

if "EXEC_KEY" in action_df.columns:
    print("EXEC_KEY present in action_df")
else:
    print("EXEC_KEY missing in action_df")
    
print("Access rows:", len(access_df))
print("Installation rows:", len(inst_df))
print("Winback rows:", len(winback_df))
print("Combined rows:", len(combined_df))
print("Actionable rows:", len(action_df))

# --------------------------------------------------
# Period Helpers - Slab Based
# P1 = 1st to 10th
# P2 = 11th to 20th
# P3 = 21st to month end
# --------------------------------------------------
def get_latest_month_start(df, date_col="INSTALLATION DATE"):
    valid = df.dropna(subset=[date_col]).copy()
    if valid.empty:
        return None
    return valid[date_col].max().to_period("M").to_timestamp()

def get_previous_month_start(month_start):
    if month_start is None or pd.isna(month_start):
        return None
    return (pd.Timestamp(month_start) - pd.DateOffset(months=1)).to_period("M").to_timestamp()

def get_month_end(dt):
    last_day = calendar.monthrange(dt.year, dt.month)[1]
    return pd.Timestamp(dt.year, dt.month, last_day)

def get_period_slab(dt):
    if dt.day <= 10:
        return "P1"
    elif dt.day <= 20:
        return "P2"
    else:
        return "P3"

def get_slab_date_range(month_start, slab):
    month_start = pd.Timestamp(month_start)
    year = month_start.year
    month = month_start.month
    month_end_day = calendar.monthrange(year, month)[1]

    if slab == "P1":
        start = pd.Timestamp(year, month, 1)
        end = pd.Timestamp(year, month, 10)
    elif slab == "P2":
        start = pd.Timestamp(year, month, 11)
        end = pd.Timestamp(year, month, 20)
    else:
        start = pd.Timestamp(year, month, 21)
        end = pd.Timestamp(year, month, month_end_day)

    return start, end

def get_reporting_periods(action_df):
    valid = action_df.dropna(subset=["INSTALLATION DATE"]).copy()
    if valid.empty:
        return {}

    latest_date = valid["INSTALLATION DATE"].max().normalize()
    latest_month_start = latest_date.to_period("M").to_timestamp()
    previous_month_start = get_previous_month_start(latest_month_start)

    current_slab = get_period_slab(latest_date)
    current_start, current_end = get_slab_date_range(latest_month_start, current_slab)

    weekly_df = valid[
        (valid["INSTALLATION DATE"] >= current_start) &
        (valid["INSTALLATION DATE"] <= current_end)
    ].copy()

    # fallback: if no rows in current slab, use latest available slab from latest available date
    if weekly_df.empty:
        fallback_date = valid["INSTALLATION DATE"].max().normalize()
        latest_month_start = fallback_date.to_period("M").to_timestamp()
        current_slab = get_period_slab(fallback_date)
        current_start, current_end = get_slab_date_range(latest_month_start, current_slab)

        weekly_df = valid[
            (valid["INSTALLATION DATE"] >= current_start) &
            (valid["INSTALLATION DATE"] <= current_end)
        ].copy()

    prev_slab_df = pd.DataFrame()
    if previous_month_start is not None:
        prev_start, prev_end = get_slab_date_range(previous_month_start, current_slab)
        prev_slab_df = valid[
            (valid["INSTALLATION DATE"] >= prev_start) &
            (valid["INSTALLATION DATE"] <= prev_end)
        ].copy()

    fortnight_label = "Fortnight-1" if latest_date.day <= 15 else "Fortnight-2"

    fortnight_df = valid[
        (valid["INSTALLATION DATE"].dt.to_period("M").dt.to_timestamp() == latest_month_start) &
        (((valid["INSTALLATION DATE"].dt.day <= 15) & (fortnight_label == "Fortnight-1")) |
         ((valid["INSTALLATION DATE"].dt.day > 15) & (fortnight_label == "Fortnight-2")))
    ].copy()

    if fortnight_df.empty:
        fortnight_df = weekly_df.copy()

    monthly_df = valid[
        valid["INSTALLATION DATE"].dt.to_period("M").dt.to_timestamp() == latest_month_start
    ].copy()

    if monthly_df.empty and previous_month_start is not None:
        monthly_df = valid[
            valid["INSTALLATION DATE"].dt.to_period("M").dt.to_timestamp() == previous_month_start
        ].copy()

    previous_month_df = pd.DataFrame()
    if previous_month_start is not None:
        previous_month_df = valid[
            valid["INSTALLATION DATE"].dt.to_period("M").dt.to_timestamp() == previous_month_start
        ].copy()

    return {
        "latest_date": latest_date,
        "latest_month_start": latest_month_start,
        "previous_month_start": previous_month_start,
        "current_slab": current_slab,
        "current_start": current_start,
        "current_end": current_end,
        "weekly_df": weekly_df,
        "fortnight_df": fortnight_df,
        "monthly_df": monthly_df,
        "previous_month_df": previous_month_df,
        "previous_slab_df": prev_slab_df
    }
# --------------------------------------------------
# Scope Filter by Role
# --------------------------------------------------
def apply_scope_filter(df, role, region_city):
    out = df.copy()

    if role == "Regional Lead":
        allowed_cities = [x.strip() for x in str(region_city).split(",") if x.strip()]
        out = out[out["EXEC_CITY_FINAL"].isin(allowed_cities)]

    return out

# --------------------------------------------------
# Executive Summary Builder
# --------------------------------------------------
def build_exec_summary(df_period):
    if df_period.empty:
        return pd.DataFrame()

    temp = df_period.copy()

    if "EXEC_KEY" not in temp.columns and "SALES CODE" in temp.columns:
        temp["EXEC_KEY"] = temp["SALES CODE"].astype(str).str.strip()

    if "Installation" not in temp.columns:
        temp["Installation"] = (temp["Record_Type"] == "Installation").astype(int)
    if "Winback" not in temp.columns:
        temp["Winback"] = (temp["Record_Type"] == "Winback").astype(int)

    month_col = "InstallMonthYear" if "InstallMonthYear" in temp.columns else "MonthYear"

    print("\n[DEBUG] build_exec_summary columns:")
    print(temp.columns.tolist())
    print("[DEBUG] Using month column:", month_col)

    exec_summary = (
        temp.groupby(
            ["EXEC_KEY", "EXEC_NAME_FINAL", "EXEC_CITY_FINAL", "Exec_Ageing_Bucket", "Target_Avg_Month"],
            dropna=False
        )
        .agg(
            Installation=("Installation", "sum"),
            Winback=("Winback", "sum"),
            Total=("ACCOUNT NO", "count"),
            Value=("Plan Value", "sum"),
            Months=(month_col, "nunique")
        )
        .reset_index()
    )

    exec_summary["Avg_Month_Count"] = exec_summary.apply(
        lambda row: round(safe_divide(row["Total"], row["Months"]), 1), axis=1
    )
    exec_summary["Avg_Month_Value"] = exec_summary.apply(
        lambda row: round(safe_divide(row["Value"], row["Months"]), 0), axis=1
    )
    exec_summary["Avg_Plan_Value_Per_Sale"] = exec_summary.apply(
        lambda row: round(safe_divide(row["Value"], row["Total"]), 0), axis=1
    )
    exec_summary["Expectation_Status"] = exec_summary.apply(
        lambda row: expectation_label(row["Avg_Month_Count"], row["Target_Avg_Month"]),
        axis=1
    )
    exec_summary["Target_Gap"] = (exec_summary["Avg_Month_Count"] - exec_summary["Target_Avg_Month"]).round(0)

    return exec_summary


# --------------------------------------------------
# Ageing Summary Builder
# --------------------------------------------------
def build_ageing_summary(exec_summary):
    if exec_summary.empty:
        return pd.DataFrame()

    ageing_summary = (
        exec_summary.groupby(["Exec_Ageing_Bucket", "Expectation_Status"])
        .agg(Employee_Count=("EXEC_KEY", "nunique"))
        .reset_index()
    )

    total_emp_by_bucket = (
        exec_summary.groupby("Exec_Ageing_Bucket")
        .agg(Total_Employees=("EXEC_KEY", "nunique"))
        .reset_index()
    )

    ageing_summary = ageing_summary.merge(total_emp_by_bucket, on="Exec_Ageing_Bucket", how="left")
    ageing_summary["Percent_Contribution"] = (
        ageing_summary["Employee_Count"] / ageing_summary["Total_Employees"] * 100
    ).round(0)

    return ageing_summary

# --------------------------------------------------
# Node Summary Builder
# --------------------------------------------------
def build_node_summary(df_period):
    if df_period.empty or "INSTALLATION NODE" not in df_period.columns:
        return pd.DataFrame()

    temp = df_period.copy()
    temp["Installation"] = (temp["Record_Type"] == "Installation").astype(int)
    temp["Winback"] = (temp["Record_Type"] == "Winback").astype(int)

    node_summary = (
        temp.groupby(["City", "INSTALLATION NODE"], dropna=False)
        .agg(
            Installation=("Installation", "sum"),
            Winback=("Winback", "sum"),
            Total=("ACCOUNT NO", "count"),
            Value=("Plan Value", "sum")
        )
        .reset_index()
    )

    node_summary["Avg_Value_Per_Sale"] = node_summary.apply(
        lambda row: round(safe_divide(row["Value"], row["Total"]), 0), axis=1
    )

    return node_summary.sort_values(["Total", "Value"], ascending=[False, False])

# --------------------------------------------------
# Plan Mix Summary Builder
# --------------------------------------------------
def build_plan_mix_summary(df_period):
    if df_period.empty:
        return pd.DataFrame()

    plan_summary = (
        df_period.groupby(["ARPU_BUCKET", "VALIDITY In Months", "SPEED (Mbps)"], dropna=False)
        .agg(
            Count=("ACCOUNT NO", "count"),
            Value=("Plan Value", "sum")
        )
        .reset_index()
    )

    plan_summary["Avg_Plan_Value_Per_Sale"] = plan_summary.apply(
        lambda row: round(safe_divide(row["Value"], row["Count"]), 0), axis=1
    )

    return plan_summary.sort_values(["Value", "Count"], ascending=[False, False])

# --------------------------------------------------
# Underperformer List Builder
# --------------------------------------------------
def build_underperformers(exec_summary):
    if exec_summary.empty:
        return pd.DataFrame()

    alerts = exec_summary[exec_summary["Expectation_Status"] == "Below Expectations"].copy()

    if alerts.empty:
        return alerts

    def recommended_action(row):
        if row["Avg_Plan_Value_Per_Sale"] < 1500:
            return "Coach on higher validity / ARPU mix"
        return "Review pipeline, territory and node productivity"

    alerts["Recommended_Action"] = alerts.apply(recommended_action, axis=1)
    return alerts.sort_values(["Target_Gap", "Total"], ascending=[True, False])

# --------------------------------------------------
# Trend Comparison Summary
# --------------------------------------------------
def build_month_comparison(current_df, previous_df):
    def summarize(df_):
        if df_.empty:
            return {
                "Installation": 0,
                "Winback": 0,
                "Total": 0,
                "Value": 0,
                "Avg_Plan_Value": 0
            }

        temp = df_.copy()
        temp["Installation"] = (temp["Record_Type"] == "Installation").astype(int)
        temp["Winback"] = (temp["Record_Type"] == "Winback").astype(int)

        installation = temp["Installation"].sum()
        winback = temp["Winback"].sum()
        total = len(temp)
        value = temp["Plan Value"].sum()
        avg_plan_value = round(safe_divide(value, total), 0)

        return {
            "Installation": installation,
            "Winback": winback,
            "Total": total,
            "Value": value,
            "Avg_Plan_Value": avg_plan_value
        }

    curr = summarize(current_df)
    prev = summarize(previous_df)

    comparison = {
        "Current": curr,
        "Previous": prev,
        "Delta_Total": curr["Total"] - prev["Total"],
        "Delta_Value": curr["Value"] - prev["Value"],
        "Delta_Avg_Plan_Value": curr["Avg_Plan_Value"] - prev["Avg_Plan_Value"]
    }

    return comparison

# --------------------------------------------------
# Narrative Builders
# --------------------------------------------------
def build_exec_narrative(exec_summary, period_label):
    if exec_summary.empty:
        return f"No active Sales executive data available for {period_label}."

    total_execs = exec_summary["EXEC_KEY"].nunique()
    below_execs = exec_summary[exec_summary["Expectation_Status"] == "Below Expectations"]["EXEC_KEY"].nunique()
    meeting_execs = exec_summary[exec_summary["Expectation_Status"] == "Meeting Expectations"]["EXEC_KEY"].nunique()
    exceptional_execs = exec_summary[exec_summary["Expectation_Status"] == "Exceptional"]["EXEC_KEY"].nunique()

    top_below_bucket = (
        exec_summary[exec_summary["Expectation_Status"] == "Below Expectations"]["Exec_Ageing_Bucket"].value_counts().idxmax()
        if below_execs > 0 else "None"
    )
    top_below_city = (
        exec_summary[exec_summary["Expectation_Status"] == "Below Expectations"]["EXEC_CITY_FINAL"].value_counts().idxmax()
        if below_execs > 0 else "None"
    )

    return (
        f"For {period_label}, there are {total_execs} active Sales executives in scope. "
        f"{below_execs} are below expectations, {meeting_execs} are on track, and {exceptional_execs} are exceptional. "
        f"The highest concentration of underperformance is in {top_below_bucket} bucket and most visible in {top_below_city} employee city."
    )

def build_node_narrative(node_summary, period_label):
    if node_summary.empty:
        return f"No node-level data available for {period_label}."

    top_node_by_count = node_summary.iloc[0]["INSTALLATION NODE"]
    top_node_by_value = node_summary.sort_values("Value", ascending=False).iloc[0]["INSTALLATION NODE"]

    return (
        f"For {period_label}, the strongest node by contribution count is {top_node_by_count}, "
        f"while the highest value concentration is on {top_node_by_value}. "
        f"Leaders should review high-count but low-value nodes for upsell opportunities."
    )

def build_plan_narrative(plan_summary, period_label):
    if plan_summary.empty:
        return f"No plan mix data available for {period_label}."

    low_yield = plan_summary[
        (plan_summary["ARPU_BUCKET"].isin(["upto 300", "301-500"])) &
        (plan_summary["VALIDITY In Months"].isin([1, 3]))
    ]["Count"].sum()

    strong_mix = plan_summary[
        (plan_summary["ARPU_BUCKET"].isin(["501-750", "751+"])) &
        (plan_summary["VALIDITY In Months"].isin([6, 12]))
    ]["Count"].sum()

    return (
        f"For {period_label}, current plan selling mix shows {int(low_yield)} counts in low-yield combinations "
        f"and {int(strong_mix)} counts in strong revenue-quality combinations. "
        f"Teams selling mostly low ARPU + short-validity plans need coaching to improve value quality per sale."
    )

def build_alert_narrative(alerts_df, period_label):
    if alerts_df.empty:
        return f"No active Sales executives are below expectations in {period_label}."

    top_gap_exec = alerts_df.sort_values("Target_Gap").iloc[0]["EXEC_NAME_FINAL"]
    top_gap_value = alerts_df.sort_values("Target_Gap").iloc[0]["Target_Gap"]
    alert_count = alerts_df["EXEC_KEY"].nunique()

    return (
        f"There are {alert_count} active Sales executives below expectations in {period_label}. "
        f"The sharpest gap is with {top_gap_exec} at {top_gap_value} below target on average monthly count. "
        f"Immediate intervention should focus on territory productivity, plan mix quality, and execution discipline."
    )

# --------------------------------------------------
# Build All Report Components for a Scoped Dataset
# --------------------------------------------------
def build_report_pack(scoped_perf_df, scoped_action_df, mode="weekly", periods_perf=None, periods_action=None):
    if periods_perf is None:
        periods_perf = get_reporting_periods(scoped_perf_df)

    if periods_action is None:
        periods_action = get_reporting_periods(scoped_action_df)

    if mode == "weekly":
        current_perf_df = periods_perf["weekly_df"].copy()
        previous_perf_df = periods_perf["previous_slab_df"].copy()

        current_action_df = periods_action["weekly_df"].copy()

        period_label = f"{periods_perf['current_slab']} ({periods_perf['current_start'].date()} to {periods_perf['current_end'].date()})"

    elif mode == "fortnightly":
        current_perf_df = periods_perf["fortnight_df"].copy()
        previous_perf_df = periods_perf["previous_month_df"].copy()

        current_action_df = periods_action["fortnight_df"].copy()

        period_label = f"{'1st-15th' if periods_perf['latest_date'].day <= 15 else '16th-Month End'} of {periods_perf['latest_month_start'].strftime('%b-%Y')}"

    else:
        current_perf_df = periods_perf["monthly_df"].copy()
        previous_perf_df = periods_perf["previous_month_df"].copy()

        current_action_df = periods_action["monthly_df"].copy()

        period_label = periods_perf["latest_month_start"].strftime("%b-%Y") if periods_perf["latest_month_start"] is not None else "Monthly"

    # PERFORMANCE summaries from full scoped performance df
    node_summary = build_node_summary(current_perf_df)
    plan_summary = build_plan_mix_summary(current_perf_df)
    month_comparison = build_month_comparison(current_perf_df, previous_perf_df)

    # ACTIONABLE summaries from active sales only
    exec_summary = build_exec_summary(current_action_df)
    ageing_summary = build_ageing_summary(exec_summary)
    underperformers = build_underperformers(exec_summary)

    narratives = {
        "exec": build_exec_narrative(exec_summary, period_label),
        "node": build_node_narrative(node_summary, period_label),
        "plan": build_plan_narrative(plan_summary, period_label),
        "alert": build_alert_narrative(underperformers, period_label)
    }

    return {
        "period_label": period_label,
        "current_perf_df": current_perf_df,
        "previous_perf_df": previous_perf_df,
        "current_action_df": current_action_df,
        "exec_summary": exec_summary,
        "ageing_summary": ageing_summary,
        "node_summary": node_summary,
        "plan_summary": plan_summary,
        "underperformers": underperformers,
        "month_comparison": month_comparison,
        "narratives": narratives
    }

# --------------------------------------------------
# HTML Helpers
# --------------------------------------------------
def df_to_html(df, max_rows=15):
    if df is None or df.empty:
        return "<p>No data available.</p>"

    return (
        df.head(max_rows)
        .fillna("")
        .to_html(index=False, border=1, justify="left")
    )

def format_exec_summary_for_mail(exec_summary):
    if exec_summary.empty:
        return exec_summary

    out = exec_summary.copy()
    return out[[
        "EXEC_NAME_FINAL", "EXEC_CITY_FINAL", "Exec_Ageing_Bucket",
        "Installation", "Winback", "Total", "Value",
        "Avg_Month_Count", "Target_Avg_Month", "Target_Gap",
        "Avg_Plan_Value_Per_Sale", "Expectation_Status"
    ]].rename(columns={
        "EXEC_NAME_FINAL": "Executive",
        "EXEC_CITY_FINAL": "Employee City",
        "Exec_Ageing_Bucket": "Ageing Bucket",
        "Value": "Total Value",
        "Avg_Month_Count": "Avg/Month Count",
        "Target_Avg_Month": "Target Avg/Month",
        "Target_Gap": "Target Gap",
        "Avg_Plan_Value_Per_Sale": "Avg Plan Value/Sale",
        "Expectation_Status": "Expectation"
    })

def format_underperformers_for_mail(alerts):
    if alerts.empty:
        return alerts

    out = alerts.copy()
    return out[[
        "EXEC_NAME_FINAL", "EXEC_CITY_FINAL", "Exec_Ageing_Bucket",
        "Installation", "Winback", "Total",
        "Avg_Month_Count", "Target_Avg_Month", "Target_Gap",
        "Avg_Plan_Value_Per_Sale", "Recommended_Action"
    ]].rename(columns={
        "EXEC_NAME_FINAL": "Executive",
        "EXEC_CITY_FINAL": "Employee City",
        "Exec_Ageing_Bucket": "Ageing Bucket",
        "Avg_Month_Count": "Avg/Month Count",
        "Target_Avg_Month": "Target Avg/Month",
        "Target_Gap": "Target Gap",
        "Avg_Plan_Value_Per_Sale": "Avg Plan Value/Sale",
        "Recommended_Action": "Recommended Action"
    })

def format_node_summary_for_mail(node_summary):
    if node_summary.empty:
        return node_summary

    out = node_summary.copy()
    return out[[
        "City", "INSTALLATION NODE", "Installation", "Winback", "Total", "Value", "Avg_Value_Per_Sale"
    ]].rename(columns={
        "INSTALLATION NODE": "Node",
        "Value": "Total Value",
        "Avg_Value_Per_Sale": "Avg Value/Sale"
    })

def format_plan_summary_for_mail(plan_summary):
    if plan_summary.empty:
        return plan_summary

    out = plan_summary.copy()
    return out[[
        "ARPU_BUCKET", "VALIDITY In Months", "SPEED (Mbps)", "Count", "Value", "Avg_Plan_Value_Per_Sale"
    ]].rename(columns={
        "ARPU_BUCKET": "ARPU Bucket",
        "Value": "Total Value",
        "Avg_Plan_Value_Per_Sale": "Avg Plan Value/Sale"
    })

# --------------------------------------------------
# KPI Summary HTML
# --------------------------------------------------
def build_kpi_html(month_comparison):
    curr = month_comparison["Current"]
    prev = month_comparison["Previous"]

    html = f"""
    <h3>Performance Snapshot</h3>
    <table border="1" cellpadding="6" cellspacing="0">
        <tr>
            <th>Metric</th>
            <th>Current Period</th>
            <th>Previous Month</th>
            <th>Delta</th>
        </tr>
        <tr>
            <td>Installation</td>
            <td>{curr['Installation']}</td>
            <td>{prev['Installation']}</td>
            <td>{curr['Installation'] - prev['Installation']}</td>
        </tr>
        <tr>
            <td>Winback</td>
            <td>{curr['Winback']}</td>
            <td>{prev['Winback']}</td>
            <td>{curr['Winback'] - prev['Winback']}</td>
        </tr>
        <tr>
            <td>Total Count</td>
            <td>{curr['Total']}</td>
            <td>{prev['Total']}</td>
            <td>{month_comparison['Delta_Total']}</td>
        </tr>
        <tr>
            <td>Total Value</td>
            <td>{format_currency(curr['Value'])}</td>
            <td>{format_currency(prev['Value'])}</td>
            <td>{format_currency(month_comparison['Delta_Value'])}</td>
        </tr>
        <tr>
            <td>Avg Plan Value/Sale</td>
            <td>{format_currency(curr['Avg_Plan_Value'])}</td>
            <td>{format_currency(prev['Avg_Plan_Value'])}</td>
            <td>{format_currency(month_comparison['Delta_Avg_Plan_Value'])}</td>
        </tr>
    </table>
    """
    return html

# --------------------------------------------------
# Email Body Builder
# --------------------------------------------------
def build_email_body(user_row, report_pack, mode):
    role = user_row["Role"]
    username = user_row["UserID"]
    period_label = report_pack["period_label"]

    exec_summary = format_exec_summary_for_mail(report_pack["exec_summary"])
    underperformers = format_underperformers_for_mail(report_pack["underperformers"])
    node_summary = format_node_summary_for_mail(report_pack["node_summary"])
    plan_summary = format_plan_summary_for_mail(report_pack["plan_summary"])
    ageing_summary = report_pack["ageing_summary"]

    kpi_html = build_kpi_html(report_pack["month_comparison"])

    html = f"""
    <html>
    <body>
        <p>Dear {username},</p>

        <p>Please find below the <b>{mode.title()}</b> Sales Performance Review for <b>{period_label}</b>.</p>

        <h2>{role} Performance Review - {period_label}</h2>

        {kpi_html}

        <h3>Executive Performance Narrative</h3>
        <p>{report_pack['narratives']['exec']}</p>

        <h3>Node Performance Narrative</h3>
        <p>{report_pack['narratives']['node']}</p>

        <h3>Plan / Revenue Quality Narrative</h3>
        <p>{report_pack['narratives']['plan']}</p>

        <h3>Alert Narrative</h3>
        <p>{report_pack['narratives']['alert']}</p>

        <h3>Executive Summary</h3>
        {df_to_html(exec_summary, max_rows=20)}

        <h3>Ageing Summary</h3>
        {df_to_html(ageing_summary, max_rows=20)}

        <h3>Top Node Summary</h3>
        {df_to_html(node_summary, max_rows=15)}

        <h3>Plan Selling Pattern Summary</h3>
        {df_to_html(plan_summary, max_rows=15)}

        <h3>Underperformer Detail</h3>
        {df_to_html(underperformers, max_rows=20)}

        <p><b>Suggested Leadership Actions</b></p>
        <ul>
            <li>Coach executives below target on validity and ARPU mix.</li>
            <li>Review employee-city productivity for persistently weak territories.</li>
            <li>Scale playbooks from exceptional performers across similar ageing buckets.</li>
            <li>Investigate high-count but low-value nodes for upsell intervention.</li>
        </ul>

        <p>Regards,<br>Sales Performance Automation</p>
    </body>
    </html>
    """
    return html

# --------------------------------------------------
# Subject Builder
# --------------------------------------------------
def build_email_subject(user_row, report_pack, mode):
    role = user_row["UserID"]
    region_city = user_row["Region/City"]
    period_label = report_pack["period_label"]

    if role == "Manager":
        scope = "ALL"
    else:
        scope = region_city

    return f"{mode.title()} Sales Performance Review | {role} | {scope} | {period_label}"

# --------------------------------------------------
# SMTP Send Function
# --------------------------------------------------

# --------------------------------------------------
# Attachment Builder
# --------------------------------------------------
def create_report_attachment(user_row, report_pack, mode):
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_path = temp_file.name
    temp_file.close()

    exec_summary = format_exec_summary_for_mail(report_pack["exec_summary"])
    ageing_summary = report_pack["ageing_summary"].copy()
    node_summary = format_node_summary_for_mail(report_pack["node_summary"])
    plan_summary = format_plan_summary_for_mail(report_pack["plan_summary"])
    underperformers = format_underperformers_for_mail(report_pack["underperformers"])

    with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
        exec_summary.to_excel(writer, index=False, sheet_name="Executive Summary")
        ageing_summary.to_excel(writer, index=False, sheet_name="Ageing Summary")
        node_summary.to_excel(writer, index=False, sheet_name="Node Summary")
        plan_summary.to_excel(writer, index=False, sheet_name="Plan Summary")
        underperformers.to_excel(writer, index=False, sheet_name="Underperformers")

    filename = f"{mode}_{user_row['Username']}_{report_pack['period_label'].replace(' ', '_').replace(':', '').replace('/', '-')}.xlsx"
    return temp_path, filename

def send_email(receiver_email, subject, html_body, attachment_path=None, attachment_name=None):
    print("\n[SMTP] Preparing message")
    print("[SMTP] Host:", SMTP_HOST)
    print("[SMTP] Port:", SMTP_PORT)
    print("[SMTP] Sender:", SENDER_EMAIL)
    print("[SMTP] Receiver:", receiver_email)

    msg = MIMEMultipart()
    msg["Subject"] = subject
    msg["From"] = SENDER_EMAIL
    msg["To"] = receiver_email

    msg.attach(MIMEText(html_body, "html"))

    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f'attachment; filename="{attachment_name}"'
            )
            msg.attach(part)

    try:
        print("[SMTP] Opening SSL connection...")
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=30) as server:
            server.set_debuglevel(1)

            print("[SMTP] Logging in...")
            server.login(SMTP_LOGIN_USER, SENDER_PASSWORD)

            print("[SMTP] Sending email...")
            server.sendmail(SENDER_EMAIL, receiver_email, msg.as_string())

            print("[SMTP] Email sent successfully")

    except Exception as e:
        print("[SMTP] ERROR:", e)
        raise

# --------------------------------------------------
# Build and Send Report to One User
# --------------------------------------------------
def send_report_to_user(user_row, perf_df, action_df, mode="weekly"):
    role = user_row["Role"]
    region_city = user_row["Region/City"]
    email = user_row["Email"]

    scoped_perf_df = apply_scope_filter(perf_df, role=role, region_city=region_city)
    scoped_action_df = apply_scope_filter(action_df, role=role, region_city=region_city)

    periods_perf = get_reporting_periods(scoped_perf_df)
    periods_action = get_reporting_periods(scoped_action_df)

    report_pack = build_report_pack(
        scoped_perf_df=scoped_perf_df,
        scoped_action_df=scoped_action_df,
        mode=mode,
        periods_perf=periods_perf,
        periods_action=periods_action
    )

    subject = build_email_subject(user_row, report_pack, mode)
    html_body = build_email_body(user_row, report_pack, mode)

    attachment_path, attachment_name = create_report_attachment(user_row, report_pack, mode)

    print(f"Sending {mode} mail to {user_row['Username']} -> {email}")
    send_email(
        email,
        subject,
        html_body,
        attachment_path=attachment_path,
        attachment_name=attachment_name
    )

    print(f"Sent successfully to {email}")

    if os.path.exists(attachment_path):
        os.remove(attachment_path)


# --------------------------------------------------
# Bulk Send Runner
# --------------------------------------------------
def run_bulk_reports(access_df, perf_df, action_df, mode="weekly", roles_to_send=None):
    users = access_df.copy()

    if roles_to_send is not None:
        users = users[users["Role"].isin(roles_to_send)].copy()

    users = users[users["Email"].notna()].copy()
    users = users[users["Email"].astype(str).str.strip() != ""].copy()

    print(f"\nStarting {mode} report run for {len(users)} users")

    for _, user_row in users.iterrows():
        try:
            send_report_to_user(user_row, perf_df, action_df, mode=mode)
        except Exception as e:
            print(f"Failed for {user_row['Username']} ({user_row['Email']}): {e}")

    print(f"\nCompleted {mode} report run")

# --------------------------------------------------
# Frequency-specific Functions
# --------------------------------------------------
def run_weekly_reports():
    run_bulk_reports(
        access_df=access_df,
        perf_df=perf_df,
        action_df=action_df,
        mode="weekly",
        roles_to_send=["Manager", "Regional Lead"]
    )

def run_fortnightly_reports():
    run_bulk_reports(
        access_df=access_df,
        perf_df=perf_df,
        action_df=action_df,
        mode="fortnightly",
        roles_to_send=["Manager", "Regional Lead"]
    )

def run_monthly_reports():
    run_bulk_reports(
        access_df=access_df,
        perf_df=perf_df,
        action_df=action_df,
        mode="monthly",
        roles_to_send=["Manager", "Regional Lead"]
    )

# --------------------------------------------------
# Manual Test / Entry Point
# --------------------------------------------------
if __name__ == "__main__":
    print("\nChoose run mode:")
    print("1 = Weekly")
    print("2 = Fortnightly")
    print("3 = Monthly")

    choice = input("Enter choice (1/2/3): ").strip()

    if choice == "1":
        run_weekly_reports()
    elif choice == "2":
        run_fortnightly_reports()
    elif choice == "3":
        run_monthly_reports()
    else:
        print("Invalid choice. No reports sent.")