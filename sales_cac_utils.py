import pandas as pd
import numpy as np

LFHV_RATES = {
    3:  {12: 700,  "13-17": 800,  "18-22": 900,  "23+": 1000},
    6:  {12: 900,  "13-17": 1000, "18-22": 1100, "23+": 1200},
    12: {12: 1100, "13-17": 1200, "18-22": 1400, "23+": 1600},
}

def _read_csv_safe(path_or_df):
    if isinstance(path_or_df, pd.DataFrame):
        return path_or_df.copy()
    try:
        return pd.read_csv(path_or_df, encoding="utf-8-sig")
    except:
        return pd.read_csv(path_or_df, encoding="latin1")

def _clean_money(series):
    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("₹", "", regex=False)
        .str.strip(),
        errors="coerce"
    )

def _normalize_city(city):
    if pd.isna(city):
        return ""
    city = str(city).strip()
    if city.lower() == "baroda":
        return "Vadodara"
    return city.title()

def _lfhv_slab(total):
    if total == 12:
        return 12
    elif 13 <= total <= 17:
        return "13-17"
    elif 18 <= total <= 22:
        return "18-22"
    elif total >= 23:
        return "23+"
    return None

def _hflv_rate(total):
    if 13 <= total <= 20:
        return 200
    elif 21 <= total <= 25:
        return 300
    elif 26 <= total <= 30:
        return 400
    elif total >= 31:
        return 500
    return 0

def _normalize_source(source):
    s = str(source).strip().upper() if pd.notna(source) else ""
    if s in ["NEW", "SALES", "INSTALL", "INSTALLATION"]:
        return "NEW"
    if s in ["WB", "WINBACK"]:
        return "WB"
    return s

def _city_in_group(city, group_list):
    city = _normalize_city(city)
    return city in [_normalize_city(x) for x in group_list]

def _match_flat(city, source, speed, plan_amount, validity):
    city = _normalize_city(city)
    source = _normalize_source(source)

    speed_num = pd.to_numeric(pd.Series([speed]), errors="coerce").iloc[0]
    plan_num = pd.to_numeric(pd.Series([plan_amount]), errors="coerce").iloc[0]
    validity_num = pd.to_numeric(pd.Series([validity]), errors="coerce").iloc[0]

    if pd.isna(speed_num) or pd.isna(plan_num) or pd.isna(validity_num):
        return None, "No flat"

    speed_num = int(round(speed_num))
    plan_num = int(round(plan_num))
    validity_num = int(round(validity_num))

    # ---------------------------------------------------
    # FIXED FLAT PAYOUT RULES
    # ---------------------------------------------------

    # Vadodara | Sales & WB | 20 Mbps | 12M | 2799 -> 400
    if city == "Vadodara" and speed_num == 20 and validity_num == 12 and plan_num == 2799:
        return 400, "Flat: Vadodara 20Mbps 12M 2799"

    # Vadodara | Sales & WB | 30 Mbps | 6M | 1798 -> 250
    if city == "Vadodara" and speed_num == 30 and validity_num == 6 and plan_num == 1798:
        return 250, "Flat: Vadodara 30Mbps 6M 1798"

    # Vadodara | Sales & WB | 30 Mbps | 12M | 3499 -> 500
    if city == "Vadodara" and speed_num == 30 and validity_num == 12 and plan_num == 3499:
        return 500, "Flat: Vadodara 30Mbps 12M 3499"

    # Surat | Sales & WB | 20 Mbps | 6M | 1599 -> 250
    if city == "Surat" and speed_num == 20 and validity_num == 6 and plan_num == 1599:
        return 250, "Flat: Surat 20Mbps 6M 1599"

    # Surat | WB | 20 Mbps | 12M | 2998 -> 400
    if city == "Surat" and source == "WB" and speed_num == 20 and validity_num == 12 and plan_num == 2998:
        return 400, "Flat: Surat WB 20Mbps 12M 2998"

    # All | Sales & WB | 20 Mbps | 12M | 3599 -> 400
    if speed_num == 20 and validity_num == 12 and plan_num == 3599:
        if city not in ["Surat", "Vadodara", "Pune", "Ahmedabad"]:
            return 400, "Flat: All city 20Mbps 12M 3599"
        # For Surat/Vadodara/Pune/Ahmedabad this may be in slab depending on specific rule below

    # ---------------------------------------------------
    # IN SLAB RULES
    # ---------------------------------------------------

    # All | Sales & WB | 30 Mbps | 3M | 1449 -> In Slab
    if speed_num == 30 and validity_num == 3 and plan_num == 1449:
        return None, "In Slab"

    # Surat | Sales & WB | 30 Mbps | 6M | 2799 -> In Slab
    if city == "Surat" and speed_num == 30 and validity_num == 6 and plan_num == 2799:
        return None, "In Slab"

    # Surat | Sales & WB | 30 Mbps | 12M | 4999 -> In Slab
    if city == "Surat" and speed_num == 30 and validity_num == 12 and plan_num == 4999:
        return None, "In Slab"

    # Surat,Vadodara,Pune,Ahmedabad | Sales & WB | 20 Mbps | 6M | 1999 -> In Slab
    group_cities = ["Surat", "Vadodara", "Pune", "Ahmedabad"]
    if _city_in_group(city, group_cities) and speed_num == 20 and validity_num == 6 and plan_num == 1999:
        return None, "In Slab"

    # Surat,Vadodara,Pune,Ahmedabad | Sales & WB | 20 Mbps | 12M | 3599 -> In Slab
    if _city_in_group(city, group_cities) and speed_num == 20 and validity_num == 12 and plan_num == 3599:
        return None, "In Slab"

    return None, "No flat"

def _parse_date_one(value):
    if pd.isna(value):
        return pd.NaT

    s = str(value).strip()
    if s == "" or s.lower() == "nan":
        return pd.NaT

    formats = ["%d-%b-%y", "%d-%m-%y", "%d/%m/%Y", "%Y-%m-%d", "%d-%b-%Y", "%d/%m/%y"]
    for fmt in formats:
        try:
            return pd.to_datetime(s, format=fmt)
        except:
            pass

    try:
        return pd.to_datetime(s, dayfirst=True, errors="coerce")
    except:
        return pd.NaT

def _prepare_txn_df(df, source_label):
    df = df.copy()
    df.columns = df.columns.str.strip()

    if "EMP Code" in df.columns:
        df["EMP Code"] = df["EMP Code"].astype(str).str.strip()

    if "DEPARTMENT" in df.columns:
        df = df[df["DEPARTMENT"].astype(str).str.strip().str.upper() == "SALES"].copy()

    if "SALES EXEC STATUS" in df.columns:
        df = df[df["SALES EXEC STATUS"].astype(str).str.strip().str.upper() == "ACTIVE"].copy()

    df["Source"] = source_label

    df["INSTALLATION DATE"] = df["INSTALLATION DATE"].apply(_parse_date_one)
    df = df[df["INSTALLATION DATE"].notna()].copy()

    df["MonthKey"] = df["INSTALLATION DATE"].apply(lambda x: f"{x.year:04d}-{x.month:02d}")
    df["MonthYear"] = df["INSTALLATION DATE"].apply(lambda x: x.strftime("%b-%Y"))

    if "Plan Value" in df.columns:
        df["Plan Value"] = _clean_money(df["Plan Value"])
    else:
        df["Plan Value"] = np.nan

    if "SPEED (Mbps)" in df.columns:
        df["SPEED (Mbps)"] = pd.to_numeric(df["SPEED (Mbps)"], errors="coerce")
    else:
        df["SPEED (Mbps)"] = np.nan

    if "VALIDITY In Months" in df.columns:
        df["VALIDITY In Months"] = pd.to_numeric(df["VALIDITY In Months"], errors="coerce")
    else:
        df["VALIDITY In Months"] = np.nan

    if "City" in df.columns:
        df["City"] = df["City"].apply(_normalize_city)
    else:
        df["City"] = ""

    return df

def calculate_payouts(sales_file, wb_file, ctc_file):
    sales_df = _read_csv_safe(sales_file)
    wb_df = _read_csv_safe(wb_file)
    ctc_df = _read_csv_safe(ctc_file)

    ctc_df.columns = ctc_df.columns.str.strip()
    ctc_df["EMP Code"] = ctc_df["EMP Code"].astype(str).str.strip()

    ctc_df = ctc_df[
        (ctc_df["DEPARTMENT"].astype(str).str.strip().str.upper() == "SALES") &
        (ctc_df["SALES EXEC STATUS"].astype(str).str.strip().str.upper() == "ACTIVE")
    ].copy()

    ctc_df["Fixed_CTC"] = _clean_money(ctc_df["Fixed_CTC"]).fillna(0)
    ctc_df["SCHEME"] = ctc_df["SCHEME"].astype(str).str.strip().str.upper()

    ctc_map = ctc_df[["EMP Code", "Name", "Fixed_CTC", "SCHEME"]].drop_duplicates()

    sales_df = _prepare_txn_df(sales_df, "New")
    wb_df = _prepare_txn_df(wb_df, "WB")

    combined = pd.concat([sales_df, wb_df], ignore_index=True)
    combined = combined.merge(ctc_map, on="EMP Code", how="inner")

    month_counts = (
        combined.groupby(["EMP Code", "MonthKey", "MonthYear"], as_index=False)
        .agg(
            Installs=("Source", lambda x: (x == "New").sum()),
            Winbacks=("Source", lambda x: (x == "WB").sum()),
            Name=("Name", "first"),
            Fixed_CTC=("Fixed_CTC", "max"),
            SCHEME=("SCHEME", "first"),
            City=("City", "first")
        )
    )
    month_counts["Total Activations"] = month_counts["Installs"] + month_counts["Winbacks"]

    combined = combined.merge(
        month_counts[["EMP Code", "MonthKey", "MonthYear", "Installs", "Winbacks", "Total Activations"]],
        on=["EMP Code", "MonthKey", "MonthYear"],
        how="left"
    )

    def calc_row_payout(row):
        installs = row["Installs"]
        total = row["Total Activations"]
        scheme = str(row["SCHEME"]).strip().upper() if pd.notna(row["SCHEME"]) else ""
        validity = row["VALIDITY In Months"]
        city = row["City"]
        speed = row["SPEED (Mbps)"]
        plan_amount = row["Plan Value"]
        source = row["Source"]

        if installs < 8:
            return pd.Series([0.0, 0.0, "Not eligible: <8 installs"])

        slab_payout = 0.0
        remark = ""

        if scheme == "LFHV":
            if total < 12:
                return pd.Series([0.0, 0.0, "LFHV not eligible: <12 activations"])
            slab = _lfhv_slab(total)
            validity_key = int(validity) if pd.notna(validity) else None
            slab_payout = LFHV_RATES.get(validity_key, {}).get(slab, 0.0)
            remark = f"LFHV slab {slab} validity {validity_key}"

        elif scheme == "HFLV":
            if total < 13:
                return pd.Series([0.0, 0.0, "HFLV not eligible: <13 activations"])
            slab_payout = float(_hflv_rate(total))
            remark = f"HFLV rate {slab_payout}"

        else:
            return pd.Series([0.0, 0.0, "Unknown scheme"])

        # 1. Flat payout override first
        flat_amt, flat_remark = _match_flat(city, source, speed, plan_amount, validity)
        if flat_amt is not None:
            return pd.Series([float(flat_amt), 0.0, flat_remark])

        # 2. WB only <2000 rule
        source_norm = _normalize_source(source)
        if source_norm == "WB" and pd.notna(plan_amount) and float(plan_amount) < 2000:
            payable = float(slab_payout) * 0.5
            pending = float(slab_payout) * 0.5
            return pd.Series([payable, pending, remark + " | WB <2000 plan 50% payable"])

        # 3. New sales normal slab
        return pd.Series([float(slab_payout), 0.0, remark])

    combined[["ROW_PAYOUT", "ROW_PENDING", "ROW_REMARK"]] = combined.apply(calc_row_payout, axis=1)

    combined["DEBUG_MATCH"] = (
        "City=" + combined["City"].astype(str)
        + " | Source=" + combined["Source"].astype(str)
        + " | Speed=" + combined["SPEED (Mbps)"].astype(str)
        + " | Validity=" + combined["VALIDITY In Months"].astype(str)
        + " | Plan=" + combined["Plan Value"].astype(str)
    )

    final = (
        combined.groupby(["EMP Code", "MonthKey", "MonthYear"], as_index=False)
        .agg(
            Name=("Name", "first"),
            City=("City", "first"),
            Installs=("Source", lambda x: (x == "New").sum()),
            Winbacks=("Source", lambda x: (x == "WB").sum()),
            Total_Activations=("Source", "count"),
            Fixed_CTC=("Fixed_CTC", "max"),
            SCHEME=("SCHEME", "first"),
            Variable_Payout=("ROW_PAYOUT", "sum"),
            Pending_Payout=("ROW_PENDING", "sum")
        )
    )

    final["Total_Payout"] = final["Fixed_CTC"] + final["Variable_Payout"]
    final["CAC"] = final["Total_Payout"] / final["Total_Activations"].replace(0, np.nan)
    final["CAC"] = final["CAC"].fillna(0)

    num_cols = [
        "Installs", "Winbacks", "Total_Activations",
        "Fixed_CTC", "Variable_Payout", "Pending_Payout",
        "Total_Payout", "CAC"
    ]
    for col in num_cols:
        final[col] = final[col].fillna(0).round(0)

    final = final.sort_values(["MonthKey", "City", "Name"]).reset_index(drop=True)
    final = final.drop(columns=["MonthKey"])

    return final, combined
