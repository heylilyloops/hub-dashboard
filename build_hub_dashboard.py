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
    {"sheet_name": "HUB ALSUT", "site": "ALSUT", "site_label": "HUB ALSUT", "bu": "AHI"},
    # {"sheet_name": "HUB BEKASI", "site": "BEKASI", "site_label": "HUB BEKASI", "bu": "HCI"},
    # tambah hub lain di sini
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
                df[col].astype(str).str.replace(",", "").str.strip(),
                errors="coerce"
            ).fillna(0)
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
    # Drop spacer rows
    df = df.dropna(subset=["Nomor LC (SI)"])
    df = df[df["Nomor LC (SI)"].astype(str).str.strip() != ""]
    print(f"  → {len(df)} rows after clean")

    # Parse date
    df["Delivery Date"] = pd.to_datetime(df["Delivery Date"], errors="coerce")
    df = df.dropna(subset=["Delivery Date"])
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
    df["Rooster_min"] = df["Rooster"].apply(parse_time_min) if "Rooster" in df.columns else None
    df["CheckInHub_min"] = df["Check In Hub"].apply(parse_time_min) if "Check In Hub" in df.columns else None

    # Jam kerja
    if "CheckInHub_min" in df.columns and "Rooster_min" in df.columns:
        df["JamKerja_min"] = (df["CheckInHub_min"] - df["Rooster_min"]).apply(
            lambda x: x + 1440 if x is not None and x < 0 else x
        )

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
        "avg_tat_min":        a("TAT_min"),
        "avg_prep_min":       a("Preparation Time_min"),
        "avg_jamkerja_min":   a("JamKerja_min"),
        "avg_rasio_bbm":      a("Rasio BBM"),
        "tol_jalur": tol_j,
    }

def build_hub_data(df, site_cfg):
    # Daily
    daily = []
    for date, grp in df.groupby("date_str"):
        r = {"date": date}
        r.update(agg_rows(grp))
        daily.append(r)

    # Monthly
    monthly = []
    for month, grp in df.groupby("month_str"):
        r = {"month": month}
        r.update(agg_rows(grp))
        monthly.append(r)

    # Nopol summary
    nopol_summary = []
    if "Nopol" in df.columns:
        for nopol, grp in df.groupby("Nopol"):
            s = agg_rows(grp)
            s["nopol"] = nopol
            s["jenis_armada"] = grp["Jenis Armada"].mode()[0] if "Jenis Armada" in grp.columns and len(grp) else "-"
            s["jalur_list"] = grp["Jalur"].value_counts().head(3).index.tolist() if "Jalur" in grp.columns else []
            nopol_summary.append(s)

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
            print(f"  ✗ Error: {e}", file=sys.stderr)

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
