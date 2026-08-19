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
from sales_cac_utils import calculate_payouts

# --------------------------------------------------
# File Paths
# --------------------------------------------------
DATA_FILE = "New Registration Report.csv"
WINBACK_FILE = "New Winback Report.csv"
ACCESS_FILE = "Access Master.xlsx"
CTC_FILE = "TSE-ACTIVE-CTC-31 Jul 2026.csv"

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

def recommendation_from_row(row):
    installs = row.get("Installs", 0)
    total = row.get("Total_Activations", 0)
    scheme = str(row.get("SCHEME", "")).upper()
    variable = row.get("Variable_Payout", 0)

    if installs < 8:
        return "Increase new installs to minimum 8."
    if scheme == "LFHV" and total < 12:
        return "Increase total activations to at least 12."
    if scheme == "HFLV" and total < 13:
        return "Increase total activations to at least 13."
    if variable <= 0:
        return "Improve plan mix toward higher-yield / flat payout eligible plans."
    return "Maintain run-rate and improve plan quality."


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

# --------------------------------------------------
# CPS / CAC Data
# --------------------------------------------------
monthly_cac, detail_rows = calculate_payouts(
    sales_file=DATA_FILE,
    wb_file=WINBACK_FILE,
    ctc_file=CTC_FILE
)

monthly_cac["City"] = monthly_cac["City"].astype(str).str.strip()
monthly_cac["MonthYear"] = monthly_cac["MonthYear"].astype(str).str.strip()
monthly_cac["Name"] = monthly_cac["Name"].astype(str).str.strip()

detail_rows["City"] = detail_rows["City"].astype(str).str.strip()
detail_rows["MonthYear"] = detail_rows["MonthYear"].astype(str).str.strip()
detail_rows["Name"] = detail_rows["Name"].astype(str).str.strip()

monthly_cac["CPS Category"] = monthly_cac["CAC"].apply(classify_cps)
monthly_cac["Incentive Status"] = monthly_cac["Variable_Payout"].apply(incentive_status)
monthly_cac["Recommendation"] = monthly_cac.apply(recommendation_from_row, axis=1)
monthly_cac["Payout_per_Activation"] = (
    monthly_cac["Variable_Payout"] / monthly_cac["Total_Activations"].replace(0, pd.NA)
).fillna(0).round(0)


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
def apply_cps_scope_filter(cps_df, role, region_city):
    out = cps_df.copy()

    if role == "Regional Lead":
        allowed_cities = [x.strip() for x in str(region_city).split(",") if x.strip()]
        out = out[out["City"].isin(allowed_cities)]

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

    ageing_order = {
        "Less Than 3 Months": 1,
        "4-12 Months": 2,
        "Above 12 Months": 3,
        "Missing": 4
    }
    expectation_order = {
        "Below Expectations": 1,
        "Meeting Expectations": 2,
        "Exceptional": 3,
        "Missing": 4
    }

    ageing_summary["Ageing_Sort"] = ageing_summary["Exec_Ageing_Bucket"].map(ageing_order)
    ageing_summary["Expectation_Sort"] = ageing_summary["Expectation_Status"].map(expectation_order)

    ageing_summary = ageing_summary.sort_values(
        ["Ageing_Sort", "Expectation_Sort"],
        ascending=[True, True]
    ).drop(columns=["Ageing_Sort", "Expectation_Sort"])

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
        f"The highest concentration of underperformance is in {top_below_bucket} bucket and most visible in {top_below_city} city."
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
        current_label = period_label

        if periods_perf["previous_month_start"] is not None:
            prev_start, prev_end = get_slab_date_range(periods_perf["previous_month_start"], periods_perf["current_slab"])
            previous_label = f"{periods_perf['current_slab']} ({prev_start.date()} to {prev_end.date()})"
        else:
            previous_label = "Previous Comparable Period"

    elif mode == "fortnightly":
        current_perf_df = periods_perf["fortnight_df"].copy()
        previous_perf_df = periods_perf["previous_month_df"].copy()

        current_action_df = periods_action["fortnight_df"].copy()

        period_label = f"{'1st-15th' if periods_perf['latest_date'].day <= 15 else '16th-Month End'} of {periods_perf['latest_month_start'].strftime('%b-%Y')}"
        current_label = period_label
        previous_label = periods_perf["previous_month_start"].strftime("%b-%y") if periods_perf["previous_month_start"] is not None else "Previous Month"

    else:
        current_perf_df = periods_perf["monthly_df"].copy()
        previous_perf_df = periods_perf["previous_month_df"].copy()

        current_action_df = periods_action["monthly_df"].copy()

        period_label = periods_perf["latest_month_start"].strftime("%b-%Y") if periods_perf["latest_month_start"] is not None else "Monthly"
        current_label = periods_perf["latest_month_start"].strftime("%b-%y") if periods_perf["latest_month_start"] is not None else "Current Month"
        previous_label = periods_perf["previous_month_start"].strftime("%b-%y") if periods_perf["previous_month_start"] is not None else "Previous Month"

    node_summary = build_node_summary(current_perf_df)
    plan_summary = build_plan_mix_summary(current_perf_df)
    month_comparison = build_month_comparison(current_perf_df, previous_perf_df)

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
        "current_label": current_label,
        "previous_label": previous_label,
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
    
def build_cps_mail_pack(cps_df):
    if cps_df.empty:
        return {
            "city_cps": pd.DataFrame(),
            "category_summary": pd.DataFrame(),
            "incentive_summary": pd.DataFrame(),
            "high_cps": pd.DataFrame(),
            "needs_improvement": pd.DataFrame(),
            "no_incentive": pd.DataFrame(),
            "below_install_gate": pd.DataFrame(),
            "below_total_gate": pd.DataFrame(),
            "low_yield": pd.DataFrame(),
            "leakage_summary": pd.DataFrame()
        }

    city_cps = cps_df.groupby("City", as_index=False).agg(
        Executives=("EMP Code", "nunique"),
        Total_Activations=("Total_Activations", "sum"),
        Total_Payout=("Total_Payout", "sum")
    )
    city_cps["CAC"] = (
        city_cps["Total_Payout"] / city_cps["Total_Activations"].replace(0, pd.NA)
    ).fillna(0).round(0)
    city_cps["CPS Category"] = city_cps["CAC"].apply(classify_cps)

    category_emp = cps_df[["EMP Code", "CPS Category"]].drop_duplicates()
    category_summary = category_emp.groupby("CPS Category", as_index=False).agg(
        Employees=("EMP Code", "nunique")
    )
    category_value = cps_df.groupby("CPS Category", as_index=False).agg(
        Total_Activations=("Total_Activations", "sum"),
        Total_Payout=("Total_Payout", "sum")
    )
    category_summary = category_summary.merge(category_value, on="CPS Category", how="left")
    category_summary["Avg CPS"] = (
        category_summary["Total_Payout"] / category_summary["Total_Activations"].replace(0, pd.NA)
    ).fillna(0).round(0)

    incentive_emp = cps_df[["EMP Code", "Incentive Status"]].drop_duplicates()
    incentive_summary = incentive_emp.groupby("Incentive Status", as_index=False).agg(
        Employees=("EMP Code", "nunique")
    )
    incentive_value = cps_df.groupby("Incentive Status", as_index=False).agg(
        Total_Activations=("Total_Activations", "sum"),
        Total_Payout=("Total_Payout", "sum")
    )
    incentive_summary = incentive_summary.merge(incentive_value, on="Incentive Status", how="left")

    high_cps = cps_df[cps_df["CAC"] > 2200].copy()
    needs_improvement = cps_df[(cps_df["CAC"] >= 1800) & (cps_df["CAC"] <= 2200)].copy()
    no_incentive = cps_df[cps_df["Variable_Payout"] <= 0].copy()
    below_install_gate = cps_df[cps_df["Installs"] < 8].copy()

    below_total_gate = cps_df[
        ((cps_df["SCHEME"].astype(str).str.upper() == "LFHV") & (cps_df["Total_Activations"] < 12)) |
        ((cps_df["SCHEME"].astype(str).str.upper() == "HFLV") & (cps_df["Total_Activations"] < 13))
    ].copy()

    low_yield = cps_df[
        (
            (
                (cps_df["SCHEME"].astype(str).str.upper() == "LFHV") &
                (cps_df["Installs"] >= 8) &
                (cps_df["Total_Activations"] >= 12)
            ) |
            (
                (cps_df["SCHEME"].astype(str).str.upper() == "HFLV") &
                (cps_df["Installs"] >= 8) &
                (cps_df["Total_Activations"] >= 13)
            )
        ) &
        (cps_df["Payout_per_Activation"] < 400)
    ].copy()

    leakage_summary = pd.DataFrame({
        "Leakage Reason": [
            "Below Install Gate",
            "Below Total Activation Gate",
            "Low Yield Plan Mix",
            "No Incentive"
        ],
        "Employees": [
            below_install_gate["EMP Code"].nunique(),
            below_total_gate["EMP Code"].nunique(),
            low_yield["EMP Code"].nunique(),
            no_incentive["EMP Code"].nunique()
        ],
        "Total Activations": [
            below_install_gate["Total_Activations"].sum() if not below_install_gate.empty else 0,
            below_total_gate["Total_Activations"].sum() if not below_total_gate.empty else 0,
            low_yield["Total_Activations"].sum() if not low_yield.empty else 0,
            no_incentive["Total_Activations"].sum() if not no_incentive.empty else 0
        ]
    })

    return {
        "city_cps": city_cps,
        "category_summary": category_summary,
        "incentive_summary": incentive_summary,
        "high_cps": high_cps,
        "needs_improvement": needs_improvement,
        "no_incentive": no_incentive,
        "below_install_gate": below_install_gate,
        "below_total_gate": below_total_gate,
        "low_yield": low_yield,
        "leakage_summary": leakage_summary
    }

# --------------------------------------------------
# HTML Helpers
# --------------------------------------------------
def df_to_html(df, max_rows=15):
    if df is None or df.empty:
        return '<p style="font-size:12px;">No data available.</p>'

    out = df.head(max_rows).copy().fillna("")

    html = out.to_html(index=False, border=0, justify="center")

    html = html.replace(
        '<table class="dataframe">',
        '<table style="border-collapse:collapse;width:auto;max-width:100%;font-size:11px;border:2px solid #c00000;table-layout:auto;">'
    )
    html = html.replace(
        '<table border="1" class="dataframe">',
        '<table style="border-collapse:collapse;width:auto;max-width:100%;font-size:11px;border:2px solid #c00000;table-layout:auto;">'
    )
    html = html.replace(
        '<thead>',
        '<thead style="background-color:#c00000;color:white;text-align:center;">'
    )
    html = html.replace(
        '<th>',
        '<th style="padding:3px 6px;border:1px solid #c00000;text-align:center;white-space:nowrap;">'
    )
    html = html.replace(
        '<td>',
        '<td style="padding:3px 6px;border:1px solid #e0e0e0;text-align:center;white-space:nowrap;">'
    )

    return html


def format_exec_summary_for_mail(exec_summary):
    if exec_summary.empty:
        return exec_summary

    out = exec_summary.copy()

    expectation_order = {
        "Below Expectations": 1,
        "Meeting Expectations": 2,
        "Exceptional": 3,
        "Missing": 4
    }
    out["Expectation_Sort"] = out["Expectation_Status"].map(expectation_order)

    out = out.sort_values(
        ["Expectation_Sort", "Target_Gap", "Total"],
        ascending=[True, True, False]
    ).drop(columns=["Expectation_Sort"])

    out["Value"] = out["Value"].apply(format_inr_0)
    out["Avg_Month_Value"] = out["Avg_Month_Value"].apply(format_inr_0)
    out["Avg_Plan_Value_Per_Sale"] = out["Avg_Plan_Value_Per_Sale"].apply(format_inr_0)
    out["Avg_Month_Count"] = out["Avg_Month_Count"].round(0).apply(format_num_0)
    out["Target_Avg_Month"] = out["Target_Avg_Month"].round(0).apply(format_num_0)
    out["Target_Gap"] = out["Target_Gap"].round(0).apply(format_num_0)

    return out[[
        "EXEC_NAME_FINAL", "EXEC_CITY_FINAL", "Exec_Ageing_Bucket",
        "Installation", "Winback", "Total", "Value",
        "Avg_Month_Count", "Target_Avg_Month", "Target_Gap",
        "Avg_Plan_Value_Per_Sale", "Expectation_Status"
    ]].rename(columns={
        "EXEC_NAME_FINAL": "Executive",
        "EXEC_CITY_FINAL": "Employee City",
        "Exec_Ageing_Bucket": "Ageing Bucket",
        "Value": "₹ Total Value",
        "Avg_Month_Count": "Avg/Month Count",
        "Target_Avg_Month": "Target Avg/Month",
        "Target_Gap": "Target Gap",
        "Avg_Plan_Value_Per_Sale": "₹ Avg Plan Value/Sale",
        "Expectation_Status": "Expectation"
    })

def format_underperformers_for_mail(alerts):
    if alerts.empty:
        return alerts

    out = alerts.copy()
    out = out.sort_values(["Target_Gap", "Total"], ascending=[True, False])

    out["Avg_Month_Count"] = out["Avg_Month_Count"].round(0).apply(format_num_0)
    out["Target_Avg_Month"] = out["Target_Avg_Month"].round(0).apply(format_num_0)
    out["Target_Gap"] = out["Target_Gap"].round(0).apply(format_num_0)
    out["Avg_Plan_Value_Per_Sale"] = out["Avg_Plan_Value_Per_Sale"].apply(format_inr_0)

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
        "Avg_Plan_Value_Per_Sale": "₹ Avg Plan Value/Sale",
        "Recommended_Action": "Recommended Action"
    })

def format_node_summary_for_mail(node_summary):
    if node_summary.empty:
        return node_summary

    out = node_summary.copy()
    out = out.sort_values(["Value", "Total"], ascending=[False, False])

    out["Value"] = out["Value"].apply(format_inr_0)
    out["Avg_Value_Per_Sale"] = out["Avg_Value_Per_Sale"].apply(format_inr_0)

    return out[[
        "City", "INSTALLATION NODE", "Installation", "Winback", "Total", "Value", "Avg_Value_Per_Sale"
    ]].rename(columns={
        "INSTALLATION NODE": "Node",
        "Value": "₹ Total Value",
        "Avg_Value_Per_Sale": "₹ Avg Value/Sale"
    })

def format_plan_summary_for_mail(plan_summary):
    if plan_summary.empty:
        return plan_summary

    out = plan_summary.copy()

    arpu_order = {
        "751+": 1,
        "501-750": 2,
        "301-500": 3,
        "upto 300": 4,
        "Missing": 5
    }

    out["ARPU_Sort"] = out["ARPU_BUCKET"].map(arpu_order)
    out = out.sort_values(["ARPU_Sort", "Value", "Count"], ascending=[True, False, False]).drop(columns=["ARPU_Sort"])

    out["Value"] = out["Value"].apply(format_inr_0)
    out["Avg_Plan_Value_Per_Sale"] = out["Avg_Plan_Value_Per_Sale"].apply(format_inr_0)

    return out[[
        "ARPU_BUCKET", "VALIDITY In Months", "SPEED (Mbps)", "Count", "Value", "Avg_Plan_Value_Per_Sale"
    ]].rename(columns={
        "ARPU_BUCKET": "ARPU Bucket",
        "Value": "₹ Total Value",
        "Avg_Plan_Value_Per_Sale": "₹ Avg Plan Value/Sale"
    })
def format_inr_0(x):
    if pd.isna(x):
        return "₹0"
    try:
        return f"₹{round(float(x), 0):,.0f}"
    except:
        return x

def format_num_0(x):
    if pd.isna(x):
        return "0"
    try:
        return f"{round(float(x), 0):,.0f}"
    except:
        return x

def color_delta_html(val):
    try:
        v = float(val)
        if v > 0:
            return f'<span style="color:green;font-weight:bold;">{v:,.0f}</span>'
        elif v < 0:
            return f'<span style="color:red;font-weight:bold;">{v:,.0f}</span>'
        else:
            return f'<span style="color:#444;">{v:,.0f}</span>'
    except:
        return str(val)

def color_text_html(text):
    text = str(text)
    if "Below" in text or "Alert" in text or "Intervention" in text:
        return f'<span style="color:red;font-weight:bold;">{text}</span>'
    elif "Exceptional" in text or "On Track" in text or "Meeting" in text:
        return f'<span style="color:green;font-weight:bold;">{text}</span>'
    return text
    
def format_cps_summary_for_mail(df):
    if df.empty:
        return df
    out = df.copy()
    for col in ["Fixed_CTC", "Variable_Payout", "Pending_Payout", "Total_Payout", "CAC", "Payout_per_Activation"]:
        if col in out.columns:
            out[col] = out[col].apply(format_inr_0)
    return out

def format_simple_cps_mail_df(df):
    if df.empty:
        return df
    out = df.copy()
    for col in out.columns:
        if col in ["Total_Payout", "CAC", "Avg CPS", "Lost Opportunity", "Payout_per_Activation"]:
            out[col] = out[col].apply(format_inr_0)
    return out

# --------------------------------------------------
# KPI Summary HTML
# --------------------------------------------------
def build_kpi_html(month_comparison, current_label="Current Period", previous_label="Previous Period"):
    curr = month_comparison["Current"]
    prev = month_comparison["Previous"]

    html = f"""
    <h3 style="color:#c00000;font-size:14px;margin-bottom:6px;">📊 Performance Snapshot</h3>
    <table style="border-collapse:collapse;width:auto;max-width:100%;font-size:11px;border:2px solid #c00000;table-layout:auto;">
        <tr style="background-color:#c00000;color:white;">
            <th style="padding:3px 6px;border:1px solid #c00000;text-align:center;white-space:nowrap;">Metric</th>
            <th style="padding:3px 6px;border:1px solid #c00000;text-align:center;white-space:nowrap;">{current_label}</th>
            <th style="padding:3px 6px;border:1px solid #c00000;text-align:center;white-space:nowrap;">{previous_label}</th>
            <th style="padding:3px 6px;border:1px solid #c00000;text-align:center;white-space:nowrap;">Delta</th>
        </tr>
        <tr>
            <td style="padding:3px 6px;border:1px solid #ddd;text-align:center;white-space:nowrap;">🔧 Installation</td>
            <td style="padding:3px 6px;border:1px solid #ddd;text-align:center;white-space:nowrap;">{format_num_0(curr['Installation'])}</td>
            <td style="padding:3px 6px;border:1px solid #ddd;text-align:center;white-space:nowrap;">{format_num_0(prev['Installation'])}</td>
            <td style="padding:3px 6px;border:1px solid #ddd;text-align:center;white-space:nowrap;">{color_delta_html(curr['Installation'] - prev['Installation'])}</td>
        </tr>
        <tr>
            <td style="padding:3px 6px;border:1px solid #ddd;text-align:center;white-space:nowrap;">🔁 Winback</td>
            <td style="padding:3px 6px;border:1px solid #ddd;text-align:center;white-space:nowrap;">{format_num_0(curr['Winback'])}</td>
            <td style="padding:3px 6px;border:1px solid #ddd;text-align:center;white-space:nowrap;">{format_num_0(prev['Winback'])}</td>
            <td style="padding:3px 6px;border:1px solid #ddd;text-align:center;white-space:nowrap;">{color_delta_html(curr['Winback'] - prev['Winback'])}</td>
        </tr>
        <tr>
            <td style="padding:3px 6px;border:1px solid #ddd;text-align:center;white-space:nowrap;">📦 Total Count</td>
            <td style="padding:3px 6px;border:1px solid #ddd;text-align:center;white-space:nowrap;">{format_num_0(curr['Total'])}</td>
            <td style="padding:3px 6px;border:1px solid #ddd;text-align:center;white-space:nowrap;">{format_num_0(prev['Total'])}</td>
            <td style="padding:3px 6px;border:1px solid #ddd;text-align:center;white-space:nowrap;">{color_delta_html(month_comparison['Delta_Total'])}</td>
        </tr>
        <tr>
            <td style="padding:3px 6px;border:1px solid #ddd;text-align:center;white-space:nowrap;">💰 Total Value</td>
            <td style="padding:3px 6px;border:1px solid #ddd;text-align:center;white-space:nowrap;">{format_inr_0(curr['Value'])}</td>
            <td style="padding:3px 6px;border:1px solid #ddd;text-align:center;white-space:nowrap;">{format_inr_0(prev['Value'])}</td>
            <td style="padding:3px 6px;border:1px solid #ddd;text-align:center;white-space:nowrap;">{color_delta_html(month_comparison['Delta_Value'])}</td>
        </tr>
        <tr>
            <td style="padding:3px 6px;border:1px solid #ddd;text-align:center;white-space:nowrap;">🏷️ Avg Plan Value/Sale</td>
            <td style="padding:3px 6px;border:1px solid #ddd;text-align:center;white-space:nowrap;">{format_inr_0(curr['Avg_Plan_Value'])}</td>
            <td style="padding:3px 6px;border:1px solid #ddd;text-align:center;white-space:nowrap;">{format_inr_0(prev['Avg_Plan_Value'])}</td>
            <td style="padding:3px 6px;border:1px solid #ddd;text-align:center;white-space:nowrap;">{color_delta_html(month_comparison['Delta_Avg_Plan_Value'])}</td>
        </tr>
    </table>
    """
    return html


# --------------------------------------------------
# Email Body Builder
# --------------------------------------------------
def build_email_body(user_row, report_pack, cps_pack, mode):
    username = user_row["UserID"]
    role = user_row["Role"]
    period_label = report_pack["period_label"]

    exec_summary = format_exec_summary_for_mail(report_pack["exec_summary"])
    underperformers = format_underperformers_for_mail(report_pack["underperformers"])
    node_summary = format_node_summary_for_mail(report_pack["node_summary"])
    plan_summary = format_plan_summary_for_mail(report_pack["plan_summary"])
    ageing_summary = report_pack["ageing_summary"]
    city_cps = format_simple_cps_mail_df(cps_pack["city_cps"])
    category_summary = format_simple_cps_mail_df(cps_pack["category_summary"])
    incentive_summary = format_simple_cps_mail_df(cps_pack["incentive_summary"])
    high_cps = format_cps_summary_for_mail(cps_pack["high_cps"])
    no_incentive = format_cps_summary_for_mail(cps_pack["no_incentive"])
    below_install_gate = format_cps_summary_for_mail(cps_pack["below_install_gate"])
    below_total_gate = format_cps_summary_for_mail(cps_pack["below_total_gate"])
    low_yield = format_cps_summary_for_mail(cps_pack["low_yield"])
    leakage_summary = format_simple_cps_mail_df(cps_pack["leakage_summary"])

    kpi_html = build_kpi_html(
        report_pack["month_comparison"],
        current_label=report_pack.get("current_label", "Current Period"),
        previous_label=report_pack.get("previous_label", "Previous Period")
    )

    html = f"""
    <html>
    <body style="font-family:Arial, sans-serif; font-size:12px; color:#222;">
        <p>Dear {username},</p>

        <p>Please find below the <b>{mode.title()}</b> Sales Performance Review for <b>{period_label}</b>.</p>

        <h2 style="color:#c00000; border-bottom:2px solid #c00000; padding-bottom:4px;">
            📬 {role} Performance Review - {period_label}
        </h2>

        {kpi_html}

        <h3 style="color:#c00000;">👔 Executive Performance Narrative</h3>
        <p>{report_pack['narratives']['exec']}</p>

        <h3 style="color:#c00000;">📍 Node Performance Narrative</h3>
        <p>{report_pack['narratives']['node']}</p>

        <h3 style="color:#c00000;">📶 Plan / Revenue Quality Narrative</h3>
        <p>{report_pack['narratives']['plan']}</p>

        <h3 style="color:#c00000;">🚨 Alert Narrative</h3>
        <p style="color:red;">{report_pack['narratives']['alert']}</p>

        <h3 style="color:#c00000;">📋 Executive Summary</h3>
        {df_to_html(exec_summary, max_rows=20)}

        <h3 style="color:#c00000;">⏳ Ageing Summary</h3>
        {df_to_html(ageing_summary, max_rows=20)}

        <h3 style="color:#c00000;">🧭 Node Summary</h3>
        {df_to_html(node_summary, max_rows=15)}

        <h3 style="color:#c00000;">📊 Plan Selling Pattern Summary</h3>
        {df_to_html(plan_summary, max_rows=15)}

        <h3 style="color:#c00000;">🚩 Underperformer Detail</h3>
        {df_to_html(underperformers, max_rows=20)}
        
        <h3 style="color:#c00000;">💸 City-wise CPS Summary</h3>
        {df_to_html(city_cps, max_rows=20)}

        <h3 style="color:#c00000;">📌 CPS Category Summary</h3>
        {df_to_html(category_summary, max_rows=10)}

        <h3 style="color:#c00000;">🎯 Incentive Qualification Summary</h3>
        {df_to_html(incentive_summary, max_rows=10)}

        <h3 style="color:#c00000;">⚠️ Alarming CPS Executives (&gt; 2200)</h3>
        {df_to_html(high_cps[[
            c for c in ["EMP Code","Name","City","MonthYear","Installs","Winbacks","Total_Activations","SCHEME","Variable_Payout","Total_Payout","CAC","Recommendation"]
            if c in high_cps.columns
        ]], max_rows=15)}

        <h3 style="color:#c00000;">🚫 Executives Not Earning Incentive</h3>
        {df_to_html(no_incentive[[
            c for c in ["EMP Code","Name","City","MonthYear","Installs","Winbacks","Total_Activations","SCHEME","CAC","Recommendation"]
            if c in no_incentive.columns
        ]], max_rows=15)}

        <h3 style="color:#c00000;">📉 Incentive Leakage Summary</h3>
        {df_to_html(leakage_summary, max_rows=10)}

        <h3 style="color:#c00000;">🔻 Below Install Gate</h3>
        {df_to_html(below_install_gate[[
            c for c in ["EMP Code","Name","City","MonthYear","Installs","Winbacks","Total_Activations","Recommendation"]
            if c in below_install_gate.columns
        ]], max_rows=15)}

        <h3 style="color:#c00000;">🔻 Below Total Activation Gate</h3>
        {df_to_html(below_total_gate[[
            c for c in ["EMP Code","Name","City","MonthYear","Installs","Winbacks","Total_Activations","Recommendation"]
            if c in below_total_gate.columns
        ]], max_rows=15)}

        <h3 style="color:#c00000;">📊 Low Yield Plan Mix</h3>
        {df_to_html(low_yield[[
            c for c in ["EMP Code","Name","City","MonthYear","Installs","Winbacks","Total_Activations","Variable_Payout","Payout_per_Activation","CAC","Recommendation"]
            if c in low_yield.columns
        ]], max_rows=15)}

        <p><b style="color:#c00000;">Recommended Leadership Actions</b></p>
        <ul style="font-size:12px;">
            <li><span style="color:red;">Coach executives below target</span> on validity and ARPU mix.</li>
            <li>Review employee-city productivity for persistently weak territories.</li>
            <li><span style="color:green;">Scale practices from exceptional performers</span> across similar ageing buckets.</li>
            <li>Investigate high-count but low-value nodes for upsell intervention.</li>
        </ul>

        <p>Regards,<br><b>Sales Performance Automation</b></p>
    </body>
    </html>
    """
    return html

# --------------------------------------------------
# Subject Builder
# --------------------------------------------------
def build_email_subject(user_row, report_pack, mode):
    role = user_row["Role"]
    region_city = user_row["Region/City"]
    period_label = report_pack["period_label"]

    scope = "ALL" if "manager" in str(role).strip().lower() else region_city

    return f"📊 {mode.title()} Sales Performance + CPS Review | {role} | {scope} | {period_label}"

# --------------------------------------------------
# SMTP Send Function
# --------------------------------------------------

# --------------------------------------------------
# Attachment Builder
# --------------------------------------------------
def create_report_attachment(user_row, report_pack, cps_pack, mode):
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_path = temp_file.name
    temp_file.close()

    exec_summary = format_exec_summary_for_mail(report_pack["exec_summary"])
    ageing_summary = report_pack["ageing_summary"].copy()
    node_summary = format_node_summary_for_mail(report_pack["node_summary"])
    plan_summary = format_plan_summary_for_mail(report_pack["plan_summary"])
    underperformers = format_underperformers_for_mail(report_pack["underperformers"])
    city_cps = cps_pack["city_cps"].copy()
    category_summary = cps_pack["category_summary"].copy()
    incentive_summary = cps_pack["incentive_summary"].copy()
    high_cps = cps_pack["high_cps"].copy()
    no_incentive = cps_pack["no_incentive"].copy()
    below_install_gate = cps_pack["below_install_gate"].copy()
    below_total_gate = cps_pack["below_total_gate"].copy()
    low_yield = cps_pack["low_yield"].copy()
    leakage_summary = cps_pack["leakage_summary"].copy()

    with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
        exec_summary.to_excel(writer, index=False, sheet_name="Executive Summary")
        ageing_summary.to_excel(writer, index=False, sheet_name="Ageing Summary")
        node_summary.to_excel(writer, index=False, sheet_name="Node Summary")
        plan_summary.to_excel(writer, index=False, sheet_name="Plan Summary")
        underperformers.to_excel(writer, index=False, sheet_name="Underperformers")
        city_cps.to_excel(writer, index=False, sheet_name="CPS City Summary")
        category_summary.to_excel(writer, index=False, sheet_name="CPS Category Summary")
        incentive_summary.to_excel(writer, index=False, sheet_name="Incentive Qualification")
        high_cps.to_excel(writer, index=False, sheet_name="High CPS")
        no_incentive.to_excel(writer, index=False, sheet_name="No Incentive")
        below_install_gate.to_excel(writer, index=False, sheet_name="Below Install Gate")
        below_total_gate.to_excel(writer, index=False, sheet_name="Below Total Gate")
        low_yield.to_excel(writer, index=False, sheet_name="Low Yield Mix")
        leakage_summary.to_excel(writer, index=False, sheet_name="Leakage Summary")
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
    scoped_cps_df = apply_cps_scope_filter(monthly_cac, role=role, region_city=region_city).copy()
     
    periods_perf = get_reporting_periods(scoped_perf_df)
    periods_action = get_reporting_periods(scoped_action_df)

    report_pack = build_report_pack(
        scoped_perf_df=scoped_perf_df,
        scoped_action_df=scoped_action_df,
        mode=mode,
        periods_perf=periods_perf,
        periods_action=periods_action
    )
    
    # Filter CPS to relevant month
    if not scoped_cps_df.empty and "MonthYear" in scoped_cps_df.columns:
        if mode == "monthly" and periods_perf["latest_month_start"] is not None:
            cps_month = periods_perf["latest_month_start"].strftime("%b-%Y")
            scoped_cps_df = scoped_cps_df[scoped_cps_df["MonthYear"] == cps_month].copy()
        elif mode == "fortnightly" and periods_perf["latest_month_start"] is not None:
            cps_month = periods_perf["latest_month_start"].strftime("%b-%Y")
            scoped_cps_df = scoped_cps_df[scoped_cps_df["MonthYear"] == cps_month].copy()
        elif mode == "weekly" and periods_perf["latest_month_start"] is not None:
            cps_month = periods_perf["latest_month_start"].strftime("%b-%Y")
            scoped_cps_df = scoped_cps_df[scoped_cps_df["MonthYear"] == cps_month].copy()

    cps_pack = build_cps_mail_pack(scoped_cps_df)
    
    subject = build_email_subject(user_row, report_pack, mode)
    html_body = build_email_body(user_row, report_pack, cps_pack, mode)

    attachment_path, attachment_name = create_report_attachment(user_row, report_pack, cps_pack, mode)

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