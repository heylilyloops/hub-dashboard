"""
build_hub_dashboard.py
Pull data dari Google Sheets Hub → generate data_hub.json → commit ke repo
"""

import json, re, os, sys
from datetime import datetime
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ══════════════════════════════════════════════
# KONFIGURASI — sesuaikan bagian ini
# ══════════════════════════════════════════════

# ID Google Spreadsheet (dari URL: /spreadsheets/d/SPREADSHEET_ID/edit)
SPREADSHEET_ID = "171uYBJ0o-blWKfnfL2g9FaWWgfBkOOIhA1QejlXqKpg"

# Mapping sheet name → site key & label
# Tambah hub baru cukup di sini
HUB_SHEETS = [
    {"sheet_name": "HUB ALSUT",      "site": "ALSUT",      "site_label": "HUB ALSUT",      "bu": "AHI"},
    {"sheet_name": "HUB MAG",        "site": "MAG",        "site_label": "HUB MAG",        "bu": "AHI"},
    {"sheet_name": "HUB GANDARIA",   "site": "GANDARIA",   "site_label": "HUB GANDARIA",   "bu": "AHI"},
    {"sheet_name": "HUB KASABLANKA", "site": "KASABLANKA", "site_label": "HUB KASABLANKA", "bu": "AHI"},
    {"sheet_name": "HUB BINTARO",    "site": "BINTARO",    "site_label": "HUB BINTARO",    "bu": "AHI"},
    {"sheet_name": "HUB CIBUBUR",    "site": "CIBUBUR",    "site_label": "HUB CIBUBUR",    "bu": "AHI"},
    {"sheet_name": "HUB TERASUTERA", "site": "TERASUTERA", "site_label": "HUB TERASUTERA", "bu": "AHI"},
    {"sheet_name": "HUB PURI",       "site": "PURI",       "site_label": "HUB PURI",       "bu": "AHI"},
    {"sheet_name": "HUB DEPOK",      "site": "DEPOK",      "site_label": "HUB DEPOK",      "bu": "AHI"},
    {"sheet_name": "HUB AYB",        "site": "AYB",        "site_label": "HUB AYB",        "bu": "AHI"},
    {"sheet_name": "HUB PASKAL",     "site": "PASKAL",     "site_label": "HUB PASKAL",     "bu": "AHI"},
    {"sheet_name": "HUB IBCC",       "site": "IBCC",       "site_label": "HUB IBCC",       "bu": "AHI"},
    {"sheet_name": "HUB PALEMBANG",  "site": "PALEMBANG",  "site_label": "HUB PALEMBANG",  "bu": "AHI"},
    {"sheet_name": "HUB SEMARANG",   "site": "SEMARANG",   "site_label": "HUB SEMARANG",   "bu": "AHI"},
    {"sheet_name": "HUB YOGYA",      "site": "YOGYA",      "site_label": "HUB YOGYA",      "bu": "AHI"},
    {"sheet_name": "HUB MALANG",     "site": "MALANG",     "site_label": "HUB MALANG",     "bu": "AHI"},
    {"sheet_name": "HUB KUTA BALI",  "site": "KUTABALI",   "site_label": "HUB KUTA BALI",  "bu": "AHI"},
    {"sheet_name": "HUB DENPASAR",   "site": "DENPASAR",   "site_label": "HUB DENPASAR",   "bu": "AHI"},
]

SERVICE_ACCOUNT_FILE = "service_account.json"  # file credentials
OUTPUT_FILE = "data_hub.json"

# ══════════════════════════════════════════════
# DATA QUALITY RULES
# ══════════════════════════════════════════════

JALUR_NORMALIZE = {
    "DEPOK": "Depok",
    "CIPUTAT": "Ciputat",
    "JAKARTA BARTA": "JAKARTA BARAT",
    "jakarta pusat": "JAKARTA PUSAT",
    "CIKARANG": "Cikarang",
}

def normalize_jalur(df):
    df["Jalur"] = df["Jalur"].replace(JALUR_NORMALIZE).fillna("-")
    return df

def parse_currency(df):
    for col in ["Total UJP", "Tol", "Parkir", "BBM"]:
        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col].astype(str).str.replace(",", "").str.replace("-", "").str.strip(),
                errors="coerce"
            ).fillna(0)
    # Numeric cols yang mungkin kosong
    for col in ["Total Drop Point", "Total DO", "DO Regular", "DO RT", "DO GRW",
                "KM Delivery", "KM Start", "KM Finish", "Rasio BBM", "Volume BBM"]:
        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col].astype(str).str.replace(",", "").str.strip(),
                errors="coerce"
            )
    return df

def parse_duration_min(s):
    try:
        p = str(s).strip().split(":")
        return round(int(p[0]) * 60 + int(p[1]) + int(p[2]) / 60, 2)
    except:
        return None

def parse_time_min(s):
    try:
        s2 = re.sub(r"(\d)(AM|PM)", r"\1 \2", str(s).strip())
        t = pd.to_datetime(s2, format="%I:%M:%S %p")
        return t.hour * 60 + t.minute + t.second / 60
    except:
        try:
            t = pd.to_datetime(str(s).strip(), format="%H:%M:%S")
            return t.hour * 60 + t.minute + t.second / 60
        except:
            return None

# ══════════════════════════════════════════════
# LOAD & CLEAN
# ══════════════════════════════════════════════

def load_sheet(gc, spreadsheet_id, sheet_name):
    sh = gc.open_by_key(spreadsheet_id)
    ws = sh.worksheet(sheet_name)
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    print(f"  [{sheet_name}] {len(df)} rows loaded")
    return df

def clean_df(df):
    # Drop spacer rows — Nomor LC harus string bermakna (lebih dari 3 karakter)
    df = df[df["Nomor LC (SI)"].astype(str).str.strip().str.len() > 3]
    print(f"  → {len(df)} rows after clean")

    # Parse date — drop baris kosong / invalid (misal "0:00:00" atau kosong)
    df = df[df["Delivery Date"].astype(str).str.strip().str.len() > 3]
    df["Delivery Date"] = pd.to_datetime(df["Delivery Date"], errors="coerce")
    df = df.dropna(subset=["Delivery Date"])
    df = df[df["Delivery Date"].dt.year > 2000]
    df = df[df["Delivery Date"] <= pd.Timestamp.today()]  # drop future dates
    df["date_str"] = df["Delivery Date"].dt.strftime("%Y-%m-%d")
    df["month_str"] = df["Delivery Date"].dt.strftime("%Y-%m")

    # Normalize
    df = normalize_jalur(df)
    df = parse_currency(df)

    # Duration fields
    for col in ["TAT", "Preparation Time", "Travel Time"]:
        if col in df.columns:
            df[col + "_min"] = df[col].apply(parse_duration_min)

    # Time fields → menit sejak tengah malam
    for col_name, attr in [
        ("Rooster", "Rooster_min"),
        ("Check Out Hub", "CheckOutHub_min"),
        ("Check In First Cust", "CheckInFirstCust_min"),
        ("Check Out Last Cust", "CheckOutLastCust_min"),
        ("Check In Hub", "CheckInHub_min"),
    ]:
        if col_name in df.columns:
            df[attr] = df[col_name].apply(parse_time_min)

    # Gap fields
    def safe_gap(a, b):
        diff = df[a] - df[b]
        return diff.apply(lambda x: x + 1440 if x is not None and not pd.isna(x) and x < 0 else x)

    if "CheckInHub_min" in df.columns and "Rooster_min" in df.columns:
        df["JamKerja_min"] = safe_gap("CheckInHub_min", "Rooster_min")
    if "CheckInFirstCust_min" in df.columns and "CheckOutHub_min" in df.columns:
        df["GapBerangkat_min"] = safe_gap("CheckInFirstCust_min", "CheckOutHub_min")
    if "CheckOutLastCust_min" in df.columns and "CheckInFirstCust_min" in df.columns:
        df["Gap1stLast_min"] = safe_gap("CheckOutLastCust_min", "CheckInFirstCust_min")
    if "CheckInHub_min" in df.columns and "CheckOutLastCust_min" in df.columns:
        df["GapPulang_min"] = safe_gap("CheckInHub_min", "CheckOutLastCust_min")

    # Time per DP
    if "Gap1stLast_min" in df.columns and "Total Drop Point" in df.columns:
        df["TimePerDP_min"] = df.apply(
            lambda r: r["Gap1stLast_min"] / r["Total Drop Point"]
            if pd.notna(r["Gap1stLast_min"]) and pd.notna(r["Total Drop Point"]) and r["Total Drop Point"] > 0
            else None, axis=1
        )

    # OTD
    if "CheckOutHub_min" in df.columns and "Rooster_min" in df.columns:
        df["OTD"] = ((df["CheckOutHub_min"] <= df["Rooster_min"]) & df["CheckOutHub_min"].notna() & df["Rooster_min"].notna()).astype(float)

    return df

# ══════════════════════════════════════════════
# AGGREGATE
# ══════════════════════════════════════════════

def agg_rows(df):
    if df.empty:
        return {}
    trips = len(df)
    def s(col): return float(df[col].sum()) if col in df.columns else 0
    def a(col): 
        if col not in df.columns: return None
        v = df[col].dropna()
        return round(float(v.mean()), 2) if len(v) else None

    tol_j = {}
    if "Tol" in df.columns and "Jalur" in df.columns:
        tol_j = df.groupby("Jalur")["Tol"].sum().astype(int).to_dict()

    return {
        "trips": trips,
        "manpower": int(df["NIK Driver"].nunique()) if "NIK Driver" in df.columns else 0,
        "avg_dp": a("Total Drop Point"),
        "avg_do": a("Total DO"),
        "do_regular": int(s("DO Regular")),
        "do_rt":      int(s("DO RT")),
        "do_grw":     int(s("DO GRW")),
        "total_ujp":  int(s("Total UJP")),
        "total_tol":  int(s("Tol")),
        "total_parkir": int(s("Parkir")),
        "total_bbm":  int(s("BBM")),
        "avg_ujp":    round(float(df["Total UJP"].mean()), 0) if "Total UJP" in df.columns else None,
        "avg_km":     a("KM Delivery"),
        "avg_tat_min":          a("TAT_min"),
        "avg_prep_min":         a("Preparation Time_min"),
        "avg_travel_min":       a("Travel Time_min"),
        "avg_jamkerja_min":     a("JamKerja_min"),
        "avg_berangkat_min":    a("GapBerangkat_min"),
        "avg_1st_last_min":     a("Gap1stLast_min"),
        "avg_pulang_min":       a("GapPulang_min"),
        "avg_time_per_dp_min":  a("TimePerDP_min"),
        "avg_rasio_bbm":        a("Rasio BBM"),
        "avg_vol_bbm":          a("Volume BBM"),
        "total_vol_bbm":        round(float(df["Volume BBM"].sum()), 1) if "Volume BBM" in df.columns else 0,
        "total_km":             round(float(df["KM Delivery"].sum()), 1) if "KM Delivery" in df.columns else 0,
        "otd_count":            int(df["OTD"].sum()) if "OTD" in df.columns else 0,
        "otd_pct":              round(float(df["OTD"].sum()) / trips * 100, 1) if "OTD" in df.columns and trips else None,
        "nopol_aktif":          int(df["Nopol"].nunique()) if "Nopol" in df.columns else 0,
        "tol_jalur": tol_j,
    }

def build_hub_data(df, site_cfg):
    # Daily
    daily = []
    for date, grp in df.groupby("date_str"):
        r = {"date": date}
        r.update(agg_rows(grp))
        jalur_day = {}
        if "Jalur" in grp.columns:
            for jalur, jgrp in grp[grp["Jalur"] != "-"].groupby("Jalur"):
                jalur_day[jalur] = agg_rows(jgrp)
        r["jalur"] = jalur_day
        daily.append(r)

    # Monthly
    monthly = []
    for month, grp in df.groupby("month_str"):
        r = {"month": month}
        r.update(agg_rows(grp))
        r["manpower"] = int(grp["NIK Driver"].nunique()) if "NIK Driver" in grp.columns else 0
        monthly.append(r)

    # Nopol summary
    all_dates = sorted(df["date_str"].unique())
    nopol_summary = []
    if "Nopol" in df.columns:
        for nopol, grp in df.groupby("Nopol"):
            s = agg_rows(grp)
            s["nopol"] = nopol
            s["jenis_armada"] = grp["Jenis Armada"].mode()[0] if "Jenis Armada" in grp.columns and len(grp) else "-"
            s["jalur_list"] = grp["Jalur"].value_counts().head(3).index.tolist() if "Jalur" in grp.columns else []
            s["hari_aktif"] = int(grp["date_str"].nunique())
            s["total_hari"] = len(all_dates)
            s["utilisasi_pct"] = round(grp["date_str"].nunique() / len(all_dates) * 100, 1) if all_dates else None
            nopol_summary.append(s)

    # Jalur summary (all-time fallback)
    jalur_summary = []
    if "Jalur" in df.columns:
        for jalur, jgrp in df[df["Jalur"] != "-"].groupby("Jalur"):
            r2 = {"jalur": jalur}
            r2.update(agg_rows(jgrp))
            jalur_summary.append(r2)
        jalur_summary.sort(key=lambda x: -x["trips"])

    # Tol total
    tol_total = {}
    for d in daily:
        for j, v in (d.get("tol_jalur") or {}).items():
            tol_total[j] = tol_total.get(j, 0) + v
    tol_total = dict(sorted(tol_total.items(), key=lambda x: -x[1]))

    return {
        **site_cfg,
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "date_min": df["date_str"].min(),
        "date_max": df["date_str"].max(),
        "total_trips": len(df),
        "all_jalur": sorted(df["Jalur"].dropna().unique().tolist()) if "Jalur" in df.columns else [],
        "all_nopol": sorted(df["Nopol"].dropna().unique().tolist()) if "Nopol" in df.columns else [],
        "tol_by_jalur": tol_total,
        "nopol_summary": nopol_summary,
        "jalur_summary": jalur_summary,
        "monthly": monthly,
        "daily": daily,
    }

# ══════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════

def main():
    print("=== Hub Dashboard Build Script ===")

    # Auth
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, scope)
    gc = gspread.authorize(creds)
    print(f"Auth OK")

    hubs = []
    for cfg in HUB_SHEETS:
        print(f"\nProcessing: {cfg['sheet_name']}")
        try:
            df = load_sheet(gc, SPREADSHEET_ID, cfg["sheet_name"])
            df = clean_df(df)
            hub_data = build_hub_data(df, {
                "site": cfg["site"],
                "site_label": cfg["site_label"],
                "bu": cfg["bu"],
            })
            hubs.append(hub_data)
            print(f"  ✓ {len(hub_data['daily'])} hari, {hub_data['total_trips']} trips")
        except Exception as e:
            import traceback
            print(f"  ✗ Error: {e}")
            print(traceback.format_exc())
            continue

    output = {
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "hubs": hubs,
    }

    with open(OUTPUT_FILE, "w") as f:
        json.dump(output, f, separators=(",", ":"))

    size_kb = os.path.getsize(OUTPUT_FILE) / 1024
    print(f"\n✓ {OUTPUT_FILE} written ({size_kb:.0f}KB) — {len(hubs)} hub(s)")

if __name__ == "__main__":
    main()
