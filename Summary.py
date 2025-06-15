#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
0 Tests GUI (SQL or Excel) — SINGLE WINDOW with TWO TABS

Tabs:
  • All Summary & Data (Live)
  • DDMI Columns Compare (Live)
Features:
  - SQL or Excel source
  - Mode dropdown: Search SN / Search WO / Fetch WO
  - Single "Get Data" button per tab
  - DR8+ / FR4 device-based WO auto-fetch (via TESTRESULT_800G_MASTER.Device)
  - Fetch WOs popup with search + multi-select + "Use all"
  - Summary + Compare live views
  - PASS/FAIL colour-coding and filter dropdown (including exclude-marked variants)
  - Same WO merged in one cell in live Summary table
  - Excel export: raw sheets + formatted SUMMARY + Compare sheets
  - Marker system:
      * Manage Markers… button opens popup:
          - list markers with colour + SN count
          - add marker (name + colour)
          - assign/remove SNs
          - delete marker
      * Filter dropdown has options that exclude marked SNs
  - Splash/loading screen with smooth progress bar
"""

import os, sys, re, traceback
os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "1"
os.environ["QT_SCALE_FACTOR_ROUNDING_POLICY"] = "PassThrough"

import numpy as np
import pandas as pd
import pyodbc
import importlib.util
from trial0 import YieldFunnelWindow

from PyQt5.QtCore import (
    Qt, QSize, QThread, pyqtSignal,
    QAbstractTableModel, QModelIndex, QCoreApplication
)
from PyQt5.QtGui import QFont, QBrush, QColor
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTextEdit,
    QPushButton, QFileDialog, QMessageBox, QProgressBar,
    QTableView, QGroupBox, QListWidget, QListWidgetItem, QComboBox,
    QStyledItemDelegate, QStyleOptionViewItem, QHeaderView, QTabWidget, QMainWindow,
    QDialog, QLineEdit, QAbstractItemView, QCheckBox, QColorDialog, QAction
)
from PyQt5.QtGui import QKeySequence


# ───────────────────────────────────────────────────────────────
# Connection / tables

DB_CONN = (
    r"DRIVER=SQL Server;"
    r"SERVER=US_SQL01.USPL.HOME;"
    r"DATABASE=MES;"
    r"UID=LABVIEW;"
    r"PWD=LABVIEW;"
)

TRX_TABLE = "TestResult_800G_2XFR4_TRX_TEST"
FWWRITE_TABLE = "TestResult_800G_2XFR4_FWWRITE_TEST"
MODEHOP_TABLE = "TestResult_800G_2XFR4_TRX_Mode_Hopping"
TABLES = {
    "3TBER":   "TestResult_800G_2XFR4_Fixed_BER_Test",
    "TCBER":   "TestResult_800G_2XFR4_BER_Symbol_Error_Test",
    "SWITCH":  "TestResult_800G_2XFR4_TRX_Switch_TEST",
    "BURNIN":  "TestResult_800G_2XFR4_BURNIN_TEST",
}

MASTER_WO_TABLE = "TESTRESULT_800G_MASTER"  # COMPONENTID, WO, Device

# hard-coded Device IDs for product selection
DEVICE_ID_DR8 = "400454000023"
DEVICE_ID_FR4 = "400454000035"
#DEVICE_ID_FR4 = "400412001542" #400454000024

SUMMARY_COLUMNS = [
    ("FW Writing",   ("FWWRITE", None)),
    ("DDMI Cal",     ("TRX", "DDMI")),
    ("TP2/TP3 - RT", ("TRX", "RT")),
    ("TP2/TP3 - LT", ("TRX", "LT")),
    ("TP2/TP3 - HT", ("TRX", "HT")),
    ("Burn-in",      ("BURNIN", None)),
    ("3TBER",        ("3TBER", None)),
    ("TCBER",        ("TCBER", None)),
    ("Mode Hopping", ("MODEHOP", None)),
    ("Final Test",   ("TRX", "FINAL")),
    ("Switch Test",  ("SWITCH", None)),
]

# Station spec limits (for pivot popup colouring)
#
# Text file format (CSV, no quotes), located next to this script:
#   StationName,MetricName,Min,Max
# Example:
#   TP2/TP3 - RT,TDECQ(dB),-1.5,3.0pip
#   Burn-in,DDMI_TxP,-2.4,4.0
#

APP_DIR = os.path.dirname(os.path.abspath(__file__))
SPEC_FILE = os.path.join(APP_DIR, "spec_limits.txt")


def load_station_specs():
    specs: dict[str, dict[str, tuple[float | None, float | None]]] = {}
    if not os.path.exists(SPEC_FILE):
        return specs
    try:
        with open(SPEC_FILE, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                parts = [p.strip() for p in line.split(",")]
                if len(parts) < 4:
                    continue
                station, metric, lo_s, hi_s = parts[:4]
                if not station or not metric:
                    continue
                lo = None
                hi = None
                try:
                    if lo_s != "":
                        lo = float(lo_s)
                except Exception:
                    lo = None
                try:
                    if hi_s != "":
                        hi = float(hi_s)
                except Exception:
                    hi = None
                specs.setdefault(station, {})[metric] = (lo, hi)
    except Exception:
        # Don't break the app if file has issues
        return {}
    return specs


STATION_SPECS = load_station_specs()

# ───────────────────────────────────────────────────────────────
# Helpers (IDs, channels, gating)

def load_slot_efficiency_tab(db_conn=None):
    """
    Loads Slot efficiency.py (kept separate) and returns SlotEfficiencyTab
    so we can add it as the 4th tab.
    """
    if db_conn is None:
        db_conn = DB_CONN

    here = os.path.dirname(os.path.abspath(__file__))
    slot_path = os.path.join(here, "Slot efficiency.py")  # exact filename

    spec = importlib.util.spec_from_file_location("slot_efficiency_mod", slot_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load: {slot_path}")

    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)

    return mod.SlotEfficiencyTab(db_conn=db_conn)
SID_NUM_PAT   = re.compile(r"(\d{8,})")
STAGE_PAT     = re.compile(r"^\s*(\d+)\s*_(RT|LT|HT|ATS)\s*$", re.IGNORECASE)
BURNIN_REQUIRED_CYCLES = set(range(24))
TXP_MIN, TXP_MAX = -2.4, 4.0
RXP_MIN, RXP_MAX = -2.4, 4.0

def sid_to_number(sid) -> float | None:
    if sid is None:
        return None
    s = str(sid)
    m = SID_NUM_PAT.search(s)
    if not m:
        return None
    try:
        return float(m.group(1))
    except Exception:
        return None



def parse_lane_stage(ch: str):
    """
    Try to extract:
      - lane: a digit 1..8 anywhere in the string
      - stage: RT / LT / HT / ATS anywhere in the string
    Works for formats like '1_RT', 'CH1_RT', 'RT_CH1', '1RT', etc.
    """
    if ch is None:
        return (None, None)

    s = str(ch).strip().upper()

    # 1) First try the strict '1_RT' pattern (keeps old behaviour)
    m = STAGE_PAT.match(s)
    if m:
        lane = int(m.group(1))
        stage = m.group(2).upper()
        return (lane if 1 <= lane <= 8 else None, stage)

    # 2) Fallback: lane = any digit 1..8, stage = RT/LT/HT/ATS anywhere
    lane_match = re.search(r"(?<!\d)([1-8])(?!\d)", s)
    lane = int(lane_match.group(1)) if lane_match else None

    stage = None
    for tag in ("RT", "LT", "HT", "ATS"):
        if tag in s:
            stage = tag
            break

    if lane is None or not (1 <= lane <= 8):
        lane = None

    return (lane, stage)

def _stage_from_row_for_station(station_name: str, ch_value, stage_value) -> str:
    """
    Decide RT/LT/HT stage for one row.

    Special handling for Final Test so that:
        1_ATS     → RT
        1_ATS_LT  → LT
        1_ATS_HT  → HT
    """
    ch = "" if ch_value is None else str(ch_value).strip().upper()
    st = "" if stage_value is None else str(stage_value).strip().upper()

    # If already a clean temp stage, keep it
    if st in {"RT", "LT", "HT"}:
        return st

    # Special rules for Final Test naming with ATS suffixes
    if station_name == "Final Test":
        if "ATS_HT" in ch or ch.endswith("_HT"):
            return "HT"
        if "ATS_LT" in ch or ch.endswith("_LT"):
            return "LT"
        # plain ATS (no HT/LT) is RT
        if "ATS" in ch:
            return "RT"

    # Fallback: use the strict pattern anywhere in the string
    m = STAGE_PAT.search(ch)
    if m:
        stage = m.group(2).upper()
        if stage == "ATS" and station_name == "Final Test":
            return "RT"
        return stage

    # Last resort: if string literally contains RT/LT/HT
    for tag in ("RT", "LT", "HT"):
        if tag in ch:
            return tag

    return ""



def _find_ci_col(df: pd.DataFrame, targets: list[str]) -> str | None:
    if df is None or df.empty:
        return None
    cols = {c.lower(): c for c in df.columns}
    for t in targets:
        key = t.lower()
        if key in cols:
            return cols[key]
    return None

def _rename_burnin_channel(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    lower_map = {c.lower(): c for c in df.columns}
    if "chnumber" not in lower_map and "channel" in lower_map:
        real_col = lower_map["channel"]
        return df.rename(columns={real_col: "CHNumber"})
    return df

def normalize_test_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    out = df.copy()
    # CH_Pass_Fail
    if "CH_Pass_Fail" in out.columns:
        ser = out["CH_Pass_Fail"]
    elif "Pass/Fail" in out.columns:
        ser = out["Pass/Fail"]
    else:
        ser = ""
    out["CH_Pass_Fail"] = (
        pd.Series(ser)
        .astype(str).str.strip()
        .replace({"nan": "", "NaN": "", "None": "", "NONE": "", "NULL": "", "null": ""})
        .str.upper()
    )
    # FailureCodeID
    if "FailureCodeID" in out.columns:
        out["FailureCodeID"] = (
            out["FailureCodeID"]
            .apply(lambda x: "" if pd.isna(x) else str(x).strip())
            .replace({"nan": "", "NaN": "", "None": "", "NONE": "", "NULL": "", "null": ""})
        )
    else:
        out["FailureCodeID"] = ""
    # Lane index
    if "CHNumber" in out.columns:
        out["_CH_IDX_"] = out["CHNumber"].apply(lambda x: parse_lane_stage(x)[0])
    else:
        out["_CH_IDX_"] = None
    # Numeric SID
    if "SID" in out.columns:
        out["_SID_NUM_"] = out["SID"].apply(sid_to_number)
    else:
        out["_SID_NUM_"] = pd.NA
    return out

def keep_latest_per_channel(df: pd.DataFrame) -> pd.DataFrame:
    """
    For each (COMPONENTID, CHNumber), keep the latest SID.
    Special SID rule:
        - If both SID and SID+'A' exist for the same lane, prefer SID (no 'A').
        - If only one exists (with or without 'A'), keep whichever is present.
    """
    if df is None or df.empty:
        return df

    key_cols = [c for c in ["COMPONENTID", "CHNumber"] if c in df.columns]
    if not key_cols:
        return normalize_test_df(df)

    if "SID" in df.columns:
        sid_str = df["SID"].astype(str).str.strip()
        sidnum  = sid_str.apply(sid_to_number)

        # Priority: non-A should win if there is a tie on numeric SID
        # non-A -> 1,  A-suffix -> 0
        suffix_priority = (~sid_str.str.endswith("A")).astype(int)

        df2 = (
            df.assign(_SID_=sidnum, _SPRI_=suffix_priority)
              # sort by lane key(s), then numeric SID, then “prefer non-A”
              .sort_values(key_cols + ["_SID_", "_SPRI_"])
              # keep LAST row per lane
              .drop_duplicates(subset=key_cols, keep="last")
              .drop(columns=["_SID_", "_SPRI_"])
        )
    else:
        df2 = df.sort_values(key_cols).drop_duplicates(subset=key_cols, keep="last")

    return normalize_test_df(df2)


def _find_testnumber_col(df: pd.DataFrame) -> str | None:
    """Find the TestNumber column (case-insensitive)."""
    if df is None or df.empty:
        return None
    return _find_ci_col(df, ["TestNumber", "TESTNUMBER", "Test_Number", "TEST_NUMBER", "TestNo", "TESTNO"])


def _build_test_run_list(df: pd.DataFrame):
    """
    Return list of (display_label, testnumber_value) for the Test Number combo.
    First entry is 'Latest per CH' (None); rest are individual TestNumber values
    sorted newest-first (highest numeric value first).
    """
    items = [("Latest per CH", None)]
    if df is None or df.empty:
        return items
    tn_col = _find_testnumber_col(df)
    if tn_col is None:
        return items
    tns = df[tn_col].dropna().astype(str).str.strip().unique()
    tn_pairs = []
    for t in tns:
        if not t:
            continue
        try:
            num = float(t)
        except Exception:
            num = 0.0
        tn_pairs.append((num, t))
    tn_pairs.sort(key=lambda p: p[0], reverse=True)
    for _num, t in tn_pairs:
        items.append((t, t))
    return items


def reduce_ber_latest_per_lane_stage(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    df = normalize_test_df(df).copy()
    if "CHNumber" in df.columns:
        parsed = df["CHNumber"].apply(parse_lane_stage)
        df["_CH_IDX_"] = [p[0] for p in parsed]
        df["_STAGE_"]  = [p[1] for p in parsed]
    else:
        df["_STAGE_"] = None
    key_cols = [c for c in ["COMPONENTID", "_CH_IDX_", "_STAGE_"] if c in df.columns]
    if not key_cols:
        return df
    out = (
        df.sort_values(key_cols + ["_SID_NUM_"])
          .drop_duplicates(subset=key_cols, keep="last")
    )
    return out

def reduce_burnin_latest_per_lane_cycle(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    df = _rename_burnin_channel(df)
    df = normalize_test_df(df).copy()
    cycle_col = _find_ci_col(df, ["Test cycle", "Test_cycle", "Cycle", "TEST CYCLE"])
    if cycle_col is None:
        has_any_lane = df["_CH_IDX_"].notna().any()
        key_cols = ["COMPONENTID", "_CH_IDX_"] if has_any_lane else [c for c in ["COMPONENTID", "CHNumber"] if c in df.columns]
    else:
        key_cols = ["COMPONENTID", "_CH_IDX_", cycle_col]
    if not key_cols:
        return df
    if "_SID_NUM_" in df.columns:
        out = (
            df.sort_values(key_cols + ["_SID_NUM_"])
              .drop_duplicates(subset=key_cols, keep="last")
        )
    else:
        out = df.drop_duplicates(subset=key_cols, keep="last")
    return out



def reduce_fw_latest_per_component(df: pd.DataFrame) -> pd.DataFrame:
    """Keep a compact FW-write dataframe: latest row per COMPONENTID (best effort)."""
    if df is None or df.empty:
        return pd.DataFrame()
    d = df.copy()
    if "COMPONENTID" not in d.columns:
        for c in d.columns:
            if c.lower() == "componentid":
                d = d.rename(columns={c: "COMPONENTID"})
                break
    if "COMPONENTID" not in d.columns:
        return d
    if "Create_Time" in d.columns:
        d["_ct_"] = pd.to_datetime(d["Create_Time"], errors="coerce")
        d = d.sort_values(["COMPONENTID", "_ct_"], ascending=[True, False])
        out = d.groupby("COMPONENTID", as_index=False).head(1).drop(columns=["_ct_"], errors="ignore")
        return out.reset_index(drop=True)
    if "SID" in d.columns:
        d["_sid_"] = pd.to_numeric(d["SID"].apply(sid_to_number), errors="coerce")
        d = d.sort_values(["COMPONENTID", "_sid_"], ascending=[True, False])
        out = d.groupby("COMPONENTID", as_index=False).head(1).drop(columns=["_sid_"], errors="ignore")
        return out.reset_index(drop=True)
    return d.groupby("COMPONENTID", as_index=False).head(1).reset_index(drop=True)


def reduce_modehop_latest_testnumber(df: pd.DataFrame) -> pd.DataFrame:
    """
    Keep only the latest TestNumber per COMPONENTID, but keep ALL rows belonging
    to that latest test number (needed to evaluate overall PASS/FAIL).
    """
    if df is None or df.empty:
        return pd.DataFrame()
    d = df.copy()
    if "COMPONENTID" not in d.columns:
        for c in d.columns:
            if c.lower() == "componentid":
                d = d.rename(columns={c: "COMPONENTID"})
                break
    tn_col = _find_ci_col(d, ["TestNumber", "TESTNUMBER", "Test_Number", "TEST_NUMBER", "Test No", "TESTNO", "TestNo"])
    if tn_col is None or "COMPONENTID" not in d.columns:
        return d.reset_index(drop=True)
    d["_tn_"] = pd.to_numeric(d[tn_col], errors="coerce")
    d["COMPONENTID"] = d["COMPONENTID"].astype(str).str.strip()
    max_tn = d.groupby("COMPONENTID")["_tn_"].transform("max")
    out = d[d["_tn_"] == max_tn].drop(columns=["_tn_"], errors="ignore")
    return out.reset_index(drop=True)



def fetch_latest_modehop_testnumber_map(conn, components: list[str], log_fn=None, chunk_size: int = 800) -> dict:
    """Return mapping: COMPONENTID -> {"num": int|None, "str": str|None}

    We derive the *latest* ModeHopping TESTNUMBER per SN using MASTER, but only among TESTNUMBERs
    that actually exist in the ModeHopping table (so we don't accidentally pick a non-ModeHop run).
    """
    if not components:
        return {}

    comps = [str(x).strip() for x in components if str(x).strip()]
    if not comps:
        return {}

    cursor = conn.cursor()
    out: dict[str, dict] = {}

    for group in _chunk(comps, chunk_size):
        placeholders = ",".join(["?"] * len(group))
        sql = f"""
        WITH mh AS (
            SELECT DISTINCT
                LTRIM(RTRIM(COMPONENTID)) AS COMPONENTID,
                LTRIM(RTRIM(TESTNUMBER))  AS TESTNUMBER_STR,
                TRY_CONVERT(bigint, LTRIM(RTRIM(TESTNUMBER))) AS TESTNUMBER_NUM
            FROM {MODEHOP_TABLE}
            WHERE LTRIM(RTRIM(COMPONENTID)) IN ({placeholders})
        ),
        mm AS (
            SELECT DISTINCT
                LTRIM(RTRIM(m.COMPONENTID)) AS COMPONENTID,
                LTRIM(RTRIM(m.TESTNUMBER))  AS TESTNUMBER_STR,
                TRY_CONVERT(bigint, LTRIM(RTRIM(m.TESTNUMBER))) AS TESTNUMBER_NUM
            FROM {MASTER_WO_TABLE} m
            JOIN mh
              ON mh.COMPONENTID = LTRIM(RTRIM(m.COMPONENTID))
             AND (
                    (mh.TESTNUMBER_NUM IS NOT NULL AND TRY_CONVERT(bigint, LTRIM(RTRIM(m.TESTNUMBER))) = mh.TESTNUMBER_NUM)
                 OR (mh.TESTNUMBER_NUM IS NULL AND LTRIM(RTRIM(m.TESTNUMBER)) = mh.TESTNUMBER_STR)
                 )
        )
        SELECT
            COMPONENTID,
            MAX(TESTNUMBER_NUM) AS LatestTN_NUM,
            MAX(TESTNUMBER_STR) AS LatestTN_STR
        FROM mm
        GROUP BY COMPONENTID
        """
        try:
            cursor.execute(sql, group)
            rows = cursor.fetchall()
            for comp, latest_num, latest_str in rows:
                cid = "" if comp is None else str(comp).strip()
                if not cid:
                    continue
                out[cid] = {
                    "num": int(latest_num) if latest_num is not None else None,
                    "str": str(latest_str).strip() if latest_str is not None else None,
                }
            if log_fn:
                log_fn(f"  • ModeHop latest TESTNUMBER map: {len(rows)} SN(s) in this chunk")
        except Exception as e:
            if log_fn:
                log_fn(f"  • ModeHop latest TESTNUMBER map error: {e}")

    return out


def reduce_modehop_to_master_latest(df: pd.DataFrame, latest_map: dict) -> pd.DataFrame:
    """Filter ModeHopping rows to ONLY the MASTER-latest ModeHop TESTNUMBER per SN.

    Keeps ALL rows belonging to that latest test number (needed to evaluate overall PASS/FAIL).
    If a SN has no mapping (unexpected), we fall back to the in-table latest numeric TESTNUMBER.
    """
    if df is None or df.empty:
        return pd.DataFrame()

    d = df.copy()

    # normalize COMPONENTID
    if "COMPONENTID" not in d.columns:
        for c in d.columns:
            if c.lower() == "componentid":
                d = d.rename(columns={c: "COMPONENTID"})
                break
    if "COMPONENTID" not in d.columns:
        return d.reset_index(drop=True)

    tn_col = _find_ci_col(d, ["TestNumber", "TESTNUMBER", "Test_Number", "TEST_NUMBER", "Test No", "TESTNO", "TestNo"])
    if tn_col is None:
        return d.reset_index(drop=True)

    d["COMPONENTID"] = d["COMPONENTID"].astype(str).str.strip()
    d["_tn_str_"] = d[tn_col].astype(str).str.strip()
    d["_tn_num_"] = pd.to_numeric(d["_tn_str_"], errors="coerce")

    # mapping vectors
    num_map = {k: v.get("num") for k, v in (latest_map or {}).items()}
    str_map = {k: v.get("str") for k, v in (latest_map or {}).items()}

    d["_map_num_"] = d["COMPONENTID"].map(num_map)
    d["_map_str_"] = d["COMPONENTID"].map(str_map)

    has_map = d["_map_num_"].notna() | d["_map_str_"].notna()

    keep_map = (
        (d["_map_num_"].notna() & (d["_tn_num_"] == d["_map_num_"])) |
        (d["_map_num_"].isna() & d["_map_str_"].notna() & (d["_tn_str_"] == d["_map_str_"]))
    )

    # fallback for SNs without mapping: keep latest numeric TESTNUMBER in ModeHop table itself
    d_fallback = d[~has_map].copy()
    if not d_fallback.empty:
        d_fallback["_mx_"] = d_fallback.groupby("COMPONENTID")["_tn_num_"].transform("max")
        keep_fallback = d_fallback["_tn_num_"].notna() & (d_fallback["_tn_num_"] == d_fallback["_mx_"])
        d_fallback = d_fallback[keep_fallback].drop(columns=["_mx_"], errors="ignore")

    d_main = d[has_map & keep_map]
    out = pd.concat([d_main, d_fallback], ignore_index=True) if not d_fallback.empty else d_main

    return out.drop(columns=["_tn_str_", "_tn_num_", "_map_num_", "_map_str_"], errors="ignore").reset_index(drop=True)

def summarize_fw_write(df: pd.DataFrame, component_id: str, wo: str | None) -> str:
    """FW Writing: PASS if this ComponentID is present in FW table for this WO."""
    if df is None or df.empty:
        return ""
    cid = str(component_id).strip()
    if not cid:
        return ""
    d = df
    comp_col = "COMPONENTID" if "COMPONENTID" in d.columns else next((c for c in d.columns if c.lower()=="componentid"), None)
    if comp_col is None:
        return ""
    sub = d[d[comp_col].astype(str).str.strip() == cid]
    if sub.empty:
        return ""
    wo = "" if wo is None else str(wo).strip()
    wo_col = next((c for c in sub.columns if c.lower() in ("wo","workorder","work_order","work order","wono","wonumber")), None)
    if wo_col and wo:
        sub2 = sub[sub[wo_col].astype(str).str.strip() == wo]
        if sub2.empty:
            return ""
    return "PASS"


def summarize_mode_hopping(df: pd.DataFrame, component_id: str) -> str:
    """Mode Hopping: evaluate PASS/FAIL over ALL rows for the *latest Master TESTNUMBER*.

    IMPORTANT: QueryWorker should already reduce the ModeHopping dataframe to ONLY the latest
    TESTNUMBER per SN for the ModeHopping station (based on MASTER). This function therefore
    just checks Status across the remaining rows for that SN:
      - If ANY row is non-PASS => FAIL
      - Else PASS
    """
    if df is None or df.empty:
        return ""
    cid = str(component_id).strip()
    if not cid:
        return ""

    d = df
    comp_col = "COMPONENTID" if "COMPONENTID" in d.columns else next((c for c in d.columns if c.lower() == "componentid"), None)
    if comp_col is None:
        return ""
    sub = d[d[comp_col].astype(str).str.strip() == cid]
    if sub.empty:
        return ""

    st_col = _find_ci_col(sub, ["Status", "STATUS", "Result", "RESULT", "PassFail", "PASSFAIL"])
    if st_col is None:
        return ""

    status = sub[st_col].astype(str).str.upper().str.strip()
    return "PASS" if (status == "PASS").all() else "FAIL"

def split_trx_types(df_trx: pd.DataFrame, reduce: bool = True) -> dict:
    # Prepare empty frames for each logical stage.
    # FW Writing + Mode Hopping may not have CHNumber suffixes, so we also try to infer
    # them from Station/TestName/Type columns when present.
    result = {k: pd.DataFrame() for k in ["DDMI", "RT", "LT", "HT", "FINAL", "MODEHOP", "FW", "BURNIN"]}
    if df_trx is None or df_trx.empty:
        return result

    df = df_trx.copy()

    type_col = next((c for c in df.columns if c.lower() in ("type", "testtype", "trx_type")), None)
    station_col = next((c for c in df.columns if c.lower() in ("station", "stationname", "teststation", "test_station")), None)
    test_col = next((c for c in df.columns if c.lower() in ("testname", "test_name", "process", "operation", "opname")), None)

    ch = (
        df["CHNumber"].astype(str).str.upper()
        if "CHNumber" in df.columns
        else pd.Series("", index=df.index)
    )

    raw_src = pd.Series("", index=df.index, dtype=str)
    for c in (type_col, station_col, test_col):
        if c and c in df.columns:
            raw_src = raw_src.astype(str) + " " + df[c].astype(str)
    raw_src = raw_src.str.upper()

    # Start from original type column (if any)
    if type_col:
        tser = df[type_col].astype(str).str.upper()
    else:
        tser = pd.Series(index=df.index, dtype=object)

    # Map channels to logical stages
    tser[ch.str.endswith("_DDMI")] = "DDMI"
    tser[ch.str.endswith("_RT")]   = "RT"
    tser[ch.str.endswith("_LT")]   = "LT"
    tser[ch.str.endswith("_HT")]   = "HT"
    tser[ch.str.contains("_ATS", na=False)] = "FINAL"

    # Extra logical stages (only if not already classified as above)
    stage_keys = {"DDMI", "RT", "LT", "HT", "FINAL"}
    can_set = ~tser.fillna("").isin(list(stage_keys))

    mh_mask = raw_src.str.contains("MODE", na=False) & raw_src.str.contains("HOP", na=False)
    fw_mask = raw_src.str.contains("FW", na=False) & (raw_src.str.contains("WRIT", na=False) | raw_src.str.contains("FIRM", na=False))

    tser[mh_mask & can_set] = "MODEHOP"
    tser[fw_mask & can_set] = "FW"

    df = df.assign(_Type_=tser.fillna("TRX"))

    for key in ["DDMI", "RT", "LT", "HT", "FINAL", "MODEHOP", "FW"]:
        sub = df[df["_Type_"] == key].copy()

        # ---------- FINAL TEST SID RULE (Summary tab) ----------
        # Priority:
        #   - If both baseSID and baseSID+'A' exist -> keep only baseSID (no 'A')
        #   - If only one exists -> keep whichever is present
        if key == "FINAL" and "SID" in sub.columns:
            sid_str = sub["SID"].astype(str).str.strip()

            base_sid = sid_str.str.replace(r"A$", "", regex=True)
            ends_A = sid_str.str.endswith("A")

            tmp = pd.DataFrame({"base": base_sid, "ends_A": ends_A})
            has_nonA = tmp.groupby("base")["ends_A"].transform(lambda s: (~s).any())

            keep_mask = (has_nonA & ~ends_A) | (~has_nonA)
            sub = sub[keep_mask]

        if reduce:
            result[key] = keep_latest_per_channel(sub) if not sub.empty else pd.DataFrame()
        else:
            result[key] = normalize_test_df(sub) if not sub.empty else pd.DataFrame()

    return result




def compute_ddmi_gate_sid_for_components(trx_map: dict, components: list) -> dict:
    dd = trx_map.get("DDMI", pd.DataFrame())
    gate = {}
    if dd is None or dd.empty:
        return {sn: None for sn in components}
    for sn in components:
        sn_trim = str(sn).strip()
        sub = dd[dd["COMPONENTID"].astype(str).str.strip() == sn_trim]
        ch_set = {int(c) for c in sub["_CH_IDX_"].dropna() if 1 <= int(c) <= 8}
        if len(ch_set) < 8:
            gate[sn_trim] = None
            continue
        sid_vals = pd.to_numeric(sub["_SID_NUM_"], errors="coerce").dropna()
        gate[sn_trim] = float(sid_vals.min()) if not sid_vals.empty else None
    return gate

def summarize_one(df: pd.DataFrame, ddmi_gate_sid: float | None, apply_gate: bool) -> str:
    if df is None or df.empty:
        return ""
    df = normalize_test_df(df)
    if apply_gate:
        if ddmi_gate_sid is None:
            return ""
        latest_sid = pd.to_numeric(df["_SID_NUM_"], errors="coerce").max()
        if pd.isna(latest_sid) or latest_sid < ddmi_gate_sid:
            return ""
    lanes = {int(c) for c in df["_CH_IDX_"].dropna() if 1 <= int(c) <= 8}
    if len(lanes) < 8:
        return ""
    status = df["CH_Pass_Fail"].astype(str).str.upper().str.strip()
    if (status == "PASS").sum() == 8 and (status != "PASS").sum() == 0:
        return "PASS"
    fail_df = df[status != "PASS"].copy()
    msgs = []
    for lane, grp in fail_df.groupby("_CH_IDX_"):
        codes = grp["FailureCodeID"].astype(str).str.strip()
        codes = [c for c in codes if c not in ("", "NAN", "NONE", "NULL", "nan", "NaN", "null")]
        if not codes:
            continue
        uniq, seen = [], set()
        for c in codes:
            if c not in seen:
                seen.add(c); uniq.append(c)
        msgs.append(f"CH{int(lane)} " + "/".join(uniq))
    return ", ".join(msgs) if msgs else ""

def summarize_ber(df: pd.DataFrame, ddmi_gate_sid: float | None) -> str:
    if df is None or df.empty:
        return ""
    df = normalize_test_df(df).copy()
    if "_STAGE_" not in df.columns:
        df["_STAGE_"] = df["CHNumber"].apply(lambda x: parse_lane_stage(x)[1]) if "CHNumber" in df.columns else None
    if ddmi_gate_sid is None:
        return ""
    sid_ok = pd.to_numeric(df["_SID_NUM_"], errors="coerce") >= ddmi_gate_sid
    df = df[sid_ok]
    if df.empty:
        return ""
    lanes = sorted({int(x) for x in df["_CH_IDX_"].dropna() if 1 <= int(x) <= 8})
    if len(lanes) < 8:
        return ""
    present = (
        df.groupby(["_CH_IDX_"])["_STAGE_"]
          .apply(lambda s: set([str(x).upper() for x in s.dropna()]))
          .to_dict()
    )
    required = {"RT", "LT", "HT"}
    for lane in range(1, 9):
        if lane not in present or not required.issubset(present[lane]):
            return ""
    mask_required = df["_STAGE_"].str.upper().isin(required) & df["_CH_IDX_"].between(1, 8)
    sub = df[mask_required]
    status = sub["CH_Pass_Fail"].astype(str).str.upper().str.strip()
    if len(status) >= 24 and (status == "PASS").all():
        return "PASS"
    fail_df = sub[status != "PASS"].copy()
    msgs = []
    for lane, grp in fail_df.groupby("_CH_IDX_"):
        codes = grp["FailureCodeID"].astype(str).str.strip()
        codes = [c for c in codes if c not in ("", "NAN", "NONE", "NULL", "nan", "NaN", "null")]
        if not codes:
            continue
        uniq, seen = [], set()
        for c in codes:
            if c not in seen:
                seen.add(c); uniq.append(c)
        msgs.append(f"CH{int(lane)} " + "/".join(uniq))
    return ", ".join(msgs) if msgs else ""

def _worst_oob_with_cycle(grp: pd.DataFrame, value_col: str, cycle_col: str, lo: float, hi: float):
    best_val, best_cyc, best_dist = None, None, -1
    vals = pd.to_numeric(grp[value_col], errors="coerce")
    cycs = pd.to_numeric(grp[cycle_col], errors="coerce")
    for i, x in vals.items():
        if pd.isna(x):
            continue
        if x < lo:
            d = lo - x
        elif x > hi:
            d = x - hi
        else:
            continue
        cyc = cycs.get(i)
        try:
            cyc_int = int(cyc) if pd.notna(cyc) else None
        except Exception:
            cyc_int = None
        if d > best_dist:
            best_dist = d; best_val = float(x); best_cyc = cyc_int
    return best_val, best_cyc

def summarize_burnin(df: pd.DataFrame, ddmi_gate_sid: float | None) -> str:
    if df is None or df.empty:
        return ""
    df = _rename_burnin_channel(df)
    df = normalize_test_df(df).copy()
    cycle_col = _find_ci_col(df, ["Test cycle", "Test_cycle", "Cycle", "TEST CYCLE"])
    if cycle_col is None:
        return ""

    # ddmi_gate_sid ignored – no SID filter here

    lanes = sorted({
        int(x)
        for x in pd.to_numeric(df["_CH_IDX_"], errors="coerce").dropna()
        if 1 <= int(x) <= 8
    })
    if len(lanes) < 8:
        return ""
    by_lane_cycles = (
        df.groupby("_CH_IDX_")[cycle_col]
          .apply(lambda s: set(pd.to_numeric(s, errors="coerce").dropna().astype(int)))
          .to_dict()
    )
    for lane in range(1, 9):
        have = by_lane_cycles.get(lane, set())
        if not BURNIN_REQUIRED_CYCLES.issubset(have):
            return ""
    cycles_ok = pd.to_numeric(df[cycle_col], errors="coerce").isin(BURNIN_REQUIRED_CYCLES)
    sub = df[df["_CH_IDX_"].between(1, 8) & cycles_ok].copy()
    status = sub["CH_Pass_Fail"].astype(str).str.upper().str.strip()
    # …rest of your function (tx_col / rx_col / out_items) stays exactly as is

    if (status == "PASS").all():
        return "PASS"
    tx_col = _find_ci_col(sub, ["_TxP(dbm)DDMI", "DDMI_TxP(dBm)", "TxP", "TxPower"])
    rx_col = _find_ci_col(sub, ["DDMI_RxP(dbm)", "DDMI_RxP(dBm)", "RxP", "RxPower"])
    out_items, seen = [], set()
    fail_rows = sub[status != "PASS"].copy()
    for lane, grp in fail_rows.groupby("_CH_IDX_"):
        lane_tag = f"CH{int(lane)}"
        if tx_col:
            worst_tx, cyc_tx = _worst_oob_with_cycle(grp, tx_col, cycle_col, TXP_MIN, TXP_MAX)
            if worst_tx is not None:
                item = f"{lane_tag} DDMI TXP {worst_tx:.2f}" + (f" (cycle {cyc_tx})" if cyc_tx is not None else "")
                if item not in seen:
                    seen.add(item); out_items.append(item)
        if rx_col:
            worst_rx, cyc_rx = _worst_oob_with_cycle(grp, rx_col, cycle_col, RXP_MIN, RXP_MAX)
            if worst_rx is not None:
                item = f"{lane_tag} DDMI RXP {worst_rx:.2f}" + (f" (cycle {cyc_rx})" if cyc_rx is not None else "")
                if item not in seen:
                    seen.add(item); out_items.append(item)
    return ", ".join(out_items) if out_items else ""

# ───────────────────────────────────────────────────────────────
# Metrics / aliases

DDMI_METRICS = [
    "Power(dBm)", "DDMI_TxP", "DD_RxP1", "dTxP", "dRxP1",
    "Txp_Offset", "Rxp_Offset", "Txp_Slope", "Rxp_Slope",
    "Vcc_Slope", "DDMI_Temp", "Tops_Setpoint",
    "I2C_Vcc", "DDMI_Vcc", "dVcc(%)",
]
EXTRA_METRICS = [
    "Outer_OMA(dB)", "Outer_ER(dB)", "TDECQ(dB)", "OMA_TDECQ(dB)", "RLM",
    "TDECQ_Ceq(dB)", "Bias_Setpoint", "Tx_Disable(dBm)", "TDECQ_4.8E-6",
]
SHEET_ORDER = ["DDMI", "RT", "LT", "HT", "FINAL", "TCBER", "3TBER"]

METRIC_ALIASES = {
    "TDECQ(dB)":       ["TDECQ(dB)", "TDECQ (dB)", "TDECQ_dB"],
    "Outer_OMA(dB)":   ["Outer_OMA(dB)", "Outer OMA (dB)", "Outer_OMA_dB"],
    "Outer_ER(dB)":    ["Outer_ER(dB)", "Outer ER (dB)", "Outer_ER_dB"],
    "OMA_TDECQ(dB)":   ["OMA_TDECQ(dB)", "OMA TDECQ (dB)", "OMA_TDECQ_dB"],
    "TDECQ_Ceq(dB)":   ["TDECQ_Ceq(dB)", "TDECQ Ceq (dB)", "TDECQ_Ceq_dB"],
    "Bias_Setpoint":   ["Bias_Setpoint", "Bias Setpoint"],
    "Tx_Disable(dBm)": ["Tx_Disable(dBm)", "Tx Disable (dBm)", "Tx_Disable_dBm"],
    "TDECQ_4.8E-6":    ["TDECQ_4.8E-6", "TDECQ_4_8E-6", "TDECQ_4.8e-6"],
    "RLM":             ["RLM"],
}
def metric_targets(name: str) -> list[str]:
    return [name] + METRIC_ALIASES.get(name, [])

def _have_all_8_lanes(df: pd.DataFrame) -> bool:
    lanes = {int(x) for x in pd.to_numeric(df.get("_CH_IDX_"), errors="coerce").dropna() if 1 <= int(x) <= 8}
    return len(lanes) == 8

def _ensure_stage_col(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    if "_STAGE_" not in df.columns and "CHNumber" in df.columns:
        df = df.copy()
        df["_STAGE_"] = df["CHNumber"].apply(lambda x: parse_lane_stage(x)[1])
    return df

def _gate_apply(df: pd.DataFrame, gate_sid: float) -> pd.DataFrame:
    if df is None:
        return pd.DataFrame()
    if gate_sid is None:
        return pd.DataFrame()
    sid_ok = pd.to_numeric(df.get("_SID_NUM_"), errors="coerce") >= gate_sid
    return df[sid_ok]

def _latest_per_lane(df: pd.DataFrame, value_col: str) -> dict[int, float]:
    out = {}
    if df is None or df.empty or not value_col:
        return out
    d = df.copy()
    d["_SID_NUM_"] = pd.to_numeric(d.get("_SID_NUM_"), errors="coerce")
    d["_CH_IDX_"]  = pd.to_numeric(d.get("_CH_IDX_"), errors="coerce")
    d = d[d["_CH_IDX_"].between(1, 8)]
    vals = pd.to_numeric(d[value_col], errors="coerce")
    d = d.assign(_VAL_=vals).dropna(subset=["_VAL_"])
    if d.empty:
        return out
    idx = d.groupby("_CH_IDX_")["_SID_NUM_"].idxmax()
    picked = d.loc[idx, ["_CH_IDX_", "_VAL_"]]
    for ln, v in zip(picked["_CH_IDX_"], picked["_VAL_"]):
        try:
            ln = int(ln)
            if 1 <= ln <= 8 and pd.notna(v):
                out[ln] = float(v)
        except Exception:
            pass
    return out

# ── Profile compares for export ────────────────────────────────

def _filter_final_by_suffix(df_final: pd.DataFrame, suffix_tag: str) -> pd.DataFrame:
    df = _ensure_stage_col(normalize_test_df(df_final))
    if df is None or df.empty:
        return df
    ch = df.get("CHNumber")
    if ch is None:
        return df.iloc[0:0]
    s = ch.astype(str).str.upper()
    if suffix_tag.upper() == "ATS":
        mask = s.str.contains("_ATS")
    else:
        mask = s.str.endswith(f"_{suffix_tag.upper()}")
    return df[mask].copy()

def _values_by_sheet_profile(type_to_frames: dict, sn: str, metric: str, gate_sid: float, profile: str):
    trx_map  = type_to_frames.get("TRX", {}) if isinstance(type_to_frames.get("TRX", {}), dict) else {}
    ddmi_df  = _ensure_stage_col(normalize_test_df(trx_map.get("DDMI",  pd.DataFrame())))
    rt_df    = _ensure_stage_col(normalize_test_df(trx_map.get("RT",    pd.DataFrame())))
    lt_df    = _ensure_stage_col(normalize_test_df(trx_map.get("LT",    pd.DataFrame())))
    ht_df    = _ensure_stage_col(normalize_test_df(trx_map.get("HT",    pd.DataFrame())))
    final_df = _ensure_stage_col(normalize_test_df(trx_map.get("FINAL", pd.DataFrame())))
    final_ats_df = _ensure_stage_col(normalize_test_df(trx_map.get("FINAL_ATS", pd.DataFrame())))
    final_lt_df  = _ensure_stage_col(normalize_test_df(trx_map.get("FINAL_LT",  pd.DataFrame())))
    final_ht_df  = _ensure_stage_col(normalize_test_df(trx_map.get("FINAL_HT",  pd.DataFrame())))

    out = {}
    snk = str(sn).strip()

    if profile.upper() == "RT":
        sub = ddmi_df[ddmi_df["COMPONENTID"].astype(str).str.strip() == snk]
        if not sub.empty and _have_all_8_lanes(sub):
            col = _find_ci_col(sub, metric_targets(metric))
            if col: out["DDMI"] = _latest_per_lane(sub, col)
        sub = _gate_apply(rt_df[rt_df["COMPONENTID"].astype(str).str.strip() == snk], gate_sid)
        if not sub.empty and _have_all_8_lanes(sub):
            col = _find_ci_col(sub, metric_targets(metric))
            if col: out["RT"] = _latest_per_lane(sub, col)
        if not final_ats_df.empty:
            sub = _gate_apply(final_ats_df[final_ats_df["COMPONENTID"].astype(str).str.strip() == snk], gate_sid)
        else:
            tmp = _filter_final_by_suffix(final_df, "ATS")
            sub = _gate_apply(tmp[tmp["COMPONENTID"].astype(str).str.strip() == snk], gate_sid)
        if not sub.empty and _have_all_8_lanes(sub):
            col = _find_ci_col(sub, metric_targets(metric))
            if col: out["FINAL_ATS"] = _latest_per_lane(sub, col)

    elif profile.upper() == "LT":
        sub = _gate_apply(lt_df[lt_df["COMPONENTID"].astype(str).str.strip() == snk], gate_sid)
        if not sub.empty and _have_all_8_lanes(sub):
            col = _find_ci_col(sub, metric_targets(metric))
            if col: out["LT"] = _latest_per_lane(sub, col)
        if not final_lt_df.empty:
            sub = _gate_apply(final_lt_df[final_lt_df["COMPONENTID"].astype(str).str.strip() == snk], gate_sid)
        else:
            tmp = _filter_final_by_suffix(final_df, "LT")
            sub = _gate_apply(tmp[tmp["COMPONENTID"].astype(str).str.strip() == snk], gate_sid)
        if not sub.empty and _have_all_8_lanes(sub):
            col = _find_ci_col(sub, metric_targets(metric))
            if col: out["FINAL_LT"] = _latest_per_lane(sub, col)

    elif profile.upper() == "HT":
        sub = _gate_apply(ht_df[ht_df["COMPONENTID"].astype(str).str.strip() == snk], gate_sid)
        if not sub.empty and _have_all_8_lanes(sub):
            col = _find_ci_col(sub, metric_targets(metric))
            if col: out["HT"] = _latest_per_lane(sub, col)
        if not final_ht_df.empty:
            sub = _gate_apply(final_ht_df[final_ht_df["COMPONENTID"].astype(str).str.strip() == snk], gate_sid)
        else:
            tmp = _filter_final_by_suffix(final_df, "HT")
            sub = _gate_apply(tmp[tmp["COMPONENTID"].astype(str).str.strip() == snk], gate_sid)
        if not sub.empty and _have_all_8_lanes(sub):
            col = _find_ci_col(sub, metric_targets(metric))
            if col: out["FINAL_HT"] = _latest_per_lane(sub, col)

    return out

def build_compare_for_profile(type_to_frames: dict, components: list[str], wo_map: dict[str, str],
                              profile: str, metrics: list[str] | None = None):
    active_metrics = metrics if metrics else (DDMI_METRICS + EXTRA_METRICS)
    trx_map = type_to_frames.get("TRX", {}) if isinstance(type_to_frames.get("TRX", {}), dict) else {}
    ddmi_gate_map = compute_ddmi_gate_sid_for_components(trx_map, components)
    profile_cols = {"RT": ["DDMI", "RT", "FINAL_ATS"], "LT": ["LT", "FINAL_LT"], "HT": ["HT", "FINAL_HT"]}
    cols = profile_cols[profile.upper()]
    live_rows, blocks = [], []
    for sn in components:
        snk = str(sn).strip(); wo = wo_map.get(snk, ""); gate = ddmi_gate_map.get(snk)
        for metric in active_metrics:
            by_sheet = _values_by_sheet_profile(type_to_frames, snk, metric, gate, profile)
            if not by_sheet:
                continue
            data_rows = []
            for ch in range(1, 9):
                pool, row_vals = [], []
                for c in cols:
                    v = by_sheet.get(c, {}).get(ch, None); row_vals.append(v)
                    if v is not None: pool.append(v)
                stdac = float(pd.Series(pool, dtype=float).std(ddof=1)) if len(pool) >= 2 else (0.0 if len(pool) == 1 else None)
                data_rows.append([f"CH{ch}"] + row_vals + [stdac])
                row_payload = {"WO": wo, "SN": snk, "Metric": metric, "CH": f"CH{ch}", "StdAcross": stdac}
                for c in cols:
                    row_payload[c] = by_sheet.get(c, {}).get(ch, None)
                live_rows.append(row_payload)
            blocks.append((wo, snk, metric, pd.DataFrame(data_rows, columns=["CH"] + cols + ["StdAcross"])))
    live_cols = ["WO", "SN", "Metric", "CH"] + cols + ["StdAcross"]
    live_df = pd.DataFrame(live_rows, columns=live_cols) if live_rows else pd.DataFrame(columns=live_cols)
    return live_df.sort_values(["WO", "SN", "Metric", "CH"]).reset_index(drop=True), blocks, cols

# ── Generic compare for live tab ───────────────────────────────

def _values_by_sheet_for_sn_metric(type_to_frames: dict, sn: str, metric: str, gate_sid: float):
    trx_map = type_to_frames.get("TRX", {}) if isinstance(type_to_frames.get("TRX", {}), dict) else {}
    ddmi_df  = _ensure_stage_col(normalize_test_df(trx_map.get("DDMI",  pd.DataFrame())))
    rt_df    = _ensure_stage_col(normalize_test_df(trx_map.get("RT",    pd.DataFrame())))
    lt_df    = _ensure_stage_col(normalize_test_df(trx_map.get("LT",    pd.DataFrame())))
    ht_df    = _ensure_stage_col(normalize_test_df(trx_map.get("HT",    pd.DataFrame())))
    final_df = _ensure_stage_col(normalize_test_df(trx_map.get("FINAL", pd.DataFrame())))
    tcb_df   = _ensure_stage_col(normalize_test_df(type_to_frames.get("TCBER", pd.DataFrame())))
    thb_df   = _ensure_stage_col(normalize_test_df(type_to_frames.get("3TBER", pd.DataFrame())))

    out = {}
    snk = str(sn).strip()

    sub = ddmi_df[ddmi_df["COMPONENTID"].astype(str).str.strip() == snk]
    if not sub.empty and _have_all_8_lanes(sub):
        col = _find_ci_col(sub, metric_targets(metric))
        if col:
            out["DDMI"] = _latest_per_lane(sub, col)

    for tag, frame in (("RT", rt_df), ("LT", lt_df), ("HT", ht_df), ("FINAL", final_df)):
        sub = _gate_apply(frame[frame["COMPONENTID"].astype(str).str.strip() == snk], gate_sid)
        if not sub.empty and _have_all_8_lanes(sub):
            col = _find_ci_col(sub, metric_targets(metric))
            if col:
                out[tag] = _latest_per_lane(sub, col)

    sub = _gate_apply(tcb_df[tcb_df["COMPONENTID"].astype(str).str.strip() == snk], gate_sid)
    sub = sub[sub["_STAGE_"].astype(str).str.upper().eq("RT")]
    if not sub.empty and _have_all_8_lanes(sub):
        col = _find_ci_col(sub, metric_targets(metric))
        if col:
            out["TCBER"] = _latest_per_lane(sub, col)

    sub = _gate_apply(thb_df[thb_df["COMPONENTID"].astype(str).str.strip() == snk], gate_sid)
    sub = sub[sub["_STAGE_"].astype(str).str.upper().eq("RT")]
    if not sub.empty and _have_all_8_lanes(sub):
        col = _find_ci_col(sub, metric_targets(metric))
        if col:
            out["3TBER"] = _latest_per_lane(sub, col)

    return out

def build_ddmi_compare_live_and_blocks(type_to_frames: dict, components: list, wo_map: dict[str, str],
                                       metrics: list[str] | None = None, sheets: list[str] | None = None):
    active_metrics = metrics if metrics else (DDMI_METRICS + EXTRA_METRICS)
    active_sheets  = sheets  if sheets  else SHEET_ORDER
    trx_map = type_to_frames.get("TRX", {}) if isinstance(type_to_frames.get("TRX", {}), dict) else {}
    ddmi_gate_map = compute_ddmi_gate_sid_for_components(trx_map, components)
    live_rows, blocks = [], []
    for sn in components:
        snk = str(sn).strip(); wo = wo_map.get(snk, ""); gate = ddmi_gate_map.get(snk)
        for metric in active_metrics:
            vals_by_sheet = _values_by_sheet_for_sn_metric(type_to_frames, snk, metric, gate)
            if not vals_by_sheet:
                continue
            data_rows = []
            for ch in range(1, 9):
                pool, row_vals, col_map = [], [], {}
                for sheet in active_sheets:
                    v = vals_by_sheet.get(sheet, {}).get(ch, None)
                    row_vals.append(v); col_map[sheet] = v
                    if v is not None:
                        pool.append(v)
                stdac = float(pd.Series(pool, dtype=float).std(ddof=1)) if len(pool) >= 2 else (0.0 if len(pool) == 1 else None)
                data_rows.append([f"CH{ch}"] + row_vals + [stdac])
                row_payload = {"WO": wo, "SN": snk, "Metric": metric, "CH": f"CH{ch}", "StdAcross": stdac}
                for s in active_sheets:
                    row_payload[s] = col_map.get(s)
                live_rows.append(row_payload)
            blocks.append((wo, snk, metric, pd.DataFrame(data_rows, columns=["CH"] + active_sheets + ["StdAcross"])))
    live_cols = ["WO", "SN", "Metric", "CH"] + active_sheets + ["StdAcross"]
    live_df = pd.DataFrame(live_rows, columns=live_cols) if live_rows else pd.DataFrame(columns=live_cols)
    return live_df.sort_values(["WO", "SN", "Metric", "CH"]).reset_index(drop=True), blocks

# ───────────────────────────────────────────────────────────────
# Excel parsing → type_to_frames

def parse_excel_to_type_frames(xl_frames: dict[str, pd.DataFrame], reduce: bool = True) -> dict:
    name_map = {str(k).strip(): v for k, v in xl_frames.items()}
    low_keys = {k.lower(): k for k in name_map.keys()}

    def pick_exact(*cands):
        for c in cands:
            k = low_keys.get(c.lower())
            if k:
                return k
        return None

    def pick_regex_first(patterns):
        for k_lower in low_keys:
            for pat in patterns:
                if re.fullmatch(pat, k_lower, flags=re.IGNORECASE):
                    return low_keys[k_lower]
        return None

    _reduce_ch = keep_latest_per_channel if reduce else normalize_test_df
    _reduce_ber = reduce_ber_latest_per_lane_stage if reduce else normalize_test_df
    _reduce_bi = reduce_burnin_latest_per_lane_cycle if reduce else normalize_test_df

    out: dict = {}
    trx_sub = {}

    for sub in ["DDMI", "RT", "LT", "HT"]:
        key = pick_exact(sub)
        if key:
            df = name_map[key].copy()
            df = _reduce_ch(df)
            trx_sub[sub] = df

    key_final_rt = pick_regex_first([r"final[\s_\-]*rt"])
    key_final_lt = pick_regex_first([r"final[\s_\-]*lt"])
    key_final_ht = pick_regex_first([r"final[\s_\-]*ht"])

    parts = []
    trx_sub["FINAL_ATS"] = _reduce_ch(name_map[key_final_rt].copy()) if key_final_rt else pd.DataFrame()
    if not trx_sub["FINAL_ATS"].empty:
        parts.append(trx_sub["FINAL_ATS"])
    trx_sub["FINAL_LT"]  = _reduce_ch(name_map[key_final_lt].copy()) if key_final_lt else pd.DataFrame()
    if not trx_sub["FINAL_LT"].empty:
        parts.append(trx_sub["FINAL_LT"])
    trx_sub["FINAL_HT"]  = _reduce_ch(name_map[key_final_ht].copy()) if key_final_ht else pd.DataFrame()
    if not trx_sub["FINAL_HT"].empty:
        parts.append(trx_sub["FINAL_HT"])
    if parts:
        trx_sub["FINAL"] = _reduce_ch(pd.concat(parts, ignore_index=True))
    else:
        key_final_plain = pick_exact("FINAL")
        trx_sub["FINAL"] = _reduce_ch(name_map[key_final_plain].copy()) if key_final_plain else pd.DataFrame()

    out["TRX"] = {
        "DDMI":       trx_sub.get("DDMI", pd.DataFrame()),
        "RT":         trx_sub.get("RT", pd.DataFrame()),
        "LT":         trx_sub.get("LT", pd.DataFrame()),
        "HT":         trx_sub.get("HT", pd.DataFrame()),
        "FINAL":      trx_sub.get("FINAL", pd.DataFrame()),
        "FINAL_ATS":  trx_sub.get("FINAL_ATS", pd.DataFrame()),
        "FINAL_LT":   trx_sub.get("FINAL_LT", pd.DataFrame()),
        "FINAL_HT":   trx_sub.get("FINAL_HT", pd.DataFrame()),
    }

    def pick_any(*names):
        for n in names:
            k = low_keys.get(n.lower())
            if k:
                return k
        return None

    tcb_key = pick_any("TCBER", "BER_Symbol_Error", "TC_BER")
    out["TCBER"] = _reduce_ber(name_map[tcb_key].copy()) if tcb_key else pd.DataFrame()

    thb_key = pick_any("3TBER", "Fixed_BER", "THREE_T_BER", "3T_BER")
    out["3TBER"] = _reduce_ber(name_map[thb_key].copy()) if thb_key else pd.DataFrame()

    sw_key = pick_any("SWITCH", "TRX_SWITCH_TEST", "SWITCH_TEST")
    out["SWITCH"] = _reduce_ch(name_map[sw_key].copy()) if sw_key else pd.DataFrame()

    bi_key = pick_any("BURNIN", "BURN_IN", "BURN-IN")
    out["BURNIN"] = _reduce_bi(name_map[bi_key].copy()) if bi_key else pd.DataFrame()

    return out

# ───────────────────────────────────────────────────────────────
# WO helpers (SQL + Excel)

def _chunk(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i+n]

def fetch_table_for_components_bulk(conn, table: str, components: list[str], log_fn=None, chunk_size: int = 800) -> pd.DataFrame:
    if not components:
        return pd.DataFrame()
    cursor = conn.cursor()
    out_frames = []
    comps = [str(x).strip() for x in components if str(x).strip()]
    for group in _chunk(comps, chunk_size):
        placeholders = ",".join(["?"] * len(group))
        sql = f"SELECT * FROM {table} WHERE RTRIM(LTRIM(COMPONENTID)) IN ({placeholders})"
        try:
            cursor.execute(sql, group)
            rs = cursor.fetchall()
            if rs:
                df = pd.DataFrame.from_records(rs, columns=[c[0] for c in cursor.description])
                out_frames.append(df)
            if log_fn:
                log_fn(f"  • fetched {len(rs)} rows from {table} (batch size {len(group)})")
        except Exception as e:
            if log_fn:
                log_fn(f"  • error on {table}: {e}")
    return pd.concat(out_frames, ignore_index=True) if out_frames else pd.DataFrame()

def fetch_components_by_wos(conn, wos: list[str], log_fn=None) -> list[str]:
    if not wos:
        return []
    wos = [w.strip() for w in wos if w.strip()]
    if not wos:
        return []
    cursor = conn.cursor()
    found = set()
    for group in _chunk(wos, 800):
        placeholders = ",".join(["?"] * len(group))
        sql = f"SELECT DISTINCT RTRIM(LTRIM(COMPONENTID)) FROM {MASTER_WO_TABLE} WHERE WO IN ({placeholders})"
        try:
            cursor.execute(sql, group)
            rows = cursor.fetchall()
            for (comp,) in rows:
                if comp is not None:
                    found.add(str(comp).strip())
            if log_fn:
                log_fn(f"  • WO→SN: fetched {len(rows)} COMPONENTIDs for {len(group)} WO(s)")
        except Exception as e:
            if log_fn:
                log_fn(f"  • WO→SN lookup error: {e}")
    return sorted(found)

def fetch_wo_map(conn, components, log_fn=None, restrict_wos=None, restrict_device: str | None = None):
    if not components:
        return {}

    # Unique, trimmed SN list
    comps_unique, seen = [], set()
    for c in components:
        s = str(c).strip()
        if s and s not in seen:
            comps_unique.append(s)
            seen.add(s)

    cursor = conn.cursor()
    wo_map: dict[str, str] = {}

    # Pre-process filters (if any)
    restrict_wos = [w.strip() for w in (restrict_wos or []) if w.strip()]
    use_wo_filter = bool(restrict_wos)

    restrict_device_s = (restrict_device or "").strip()
    use_device_filter = bool(restrict_device_s)

    for group in _chunk(comps_unique, 800):
        comp_placeholders = ",".join(["?"] * len(group))

        where_parts = [f"LTRIM(RTRIM(COMPONENTID)) IN ({comp_placeholders})"]
        params = [s.strip() for s in group]

        if use_wo_filter:
            wo_placeholders = ",".join(["?"] * len(restrict_wos))
            where_parts.append(f"LTRIM(RTRIM(WO)) IN ({wo_placeholders})")
            params.extend(restrict_wos)

        if use_device_filter:
            where_parts.append("LTRIM(RTRIM(Device)) = ?")
            params.append(restrict_device_s)

        sql = (
            f"SELECT LTRIM(RTRIM(COMPONENTID)) AS COMPONENTID, "
            f"       LTRIM(RTRIM(WO))          AS WO "
            f"FROM {MASTER_WO_TABLE} "
            f"WHERE " + " AND ".join(where_parts)
        )

        try:
            cursor.execute(sql, params)
            rows = cursor.fetchall()
            for comp, wo in rows:
                comp_s = str(comp).strip()
                wo_s = "" if wo is None else str(wo).strip()
                # If multiple rows exist, keep the first seen (or change logic to prefer latest)
                if comp_s not in wo_map:
                    wo_map[comp_s] = wo_s

            if log_fn:
                extra = []
                if use_wo_filter:
                    extra.append(f"WO filter ({len(restrict_wos)})")
                if use_device_filter:
                    extra.append(f"Device={restrict_device_s}")
                extra_s = f" [{' & '.join(extra)}]" if extra else ""
                log_fn(f"  • WO lookup{extra_s}: fetched {len(rows)}/{len(group)} from {MASTER_WO_TABLE}")
        except Exception as e:
            if log_fn:
                log_fn(f"  • WO lookup error on {MASTER_WO_TABLE}: {e}")

    # Ensure all comps have a key
    for c in comps_unique:
        wo_map.setdefault(c, "")

    unresolved = [c for c in comps_unique if not wo_map.get(c)]
    if unresolved and log_fn:
        log_fn(f"  • WO missing for {len(unresolved)} SN(s): e.g. {unresolved[:5]} …")

    return wo_map



EXCEL_WO_ALIASES = ["WO", "WorkOrder", "WORKORDER", "WORK_ORDER", "WORK ORDER", "WO#", "WO_NO", "WO_NUMBER"]

def _pick_wo_column(df: pd.DataFrame) -> str | None:
    if df is None or df.empty:
        return None
    lc = {str(c).strip().lower(): c for c in df.columns}
    for cand in EXCEL_WO_ALIASES:
        if cand.lower() in lc:
            return lc[cand.lower()]
    return None

def _norm_wo_value(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if isinstance(v, (int, np.integer)):
        return str(int(v))
    if isinstance(v, (float, np.floating)):
        return str(int(v)) if float(v).is_integer() else str(v).strip()
    s = str(v).strip()
    return s[:-2] if s.endswith(".0") and s[:-2].isdigit() else s

def extract_wo_map_from_excel(frames: dict[str, pd.DataFrame], components: list[str]) -> dict[str, str]:
    targets = {str(c).strip() for c in components if str(c).strip()}
    wo_map = {c: "" for c in targets}
    priority = [
        "SUMMARY", "MASTER", "All Summary & Data (Live)", "DDMI", "RT", "LT", "HT",
        "FINAL", "FINAL RT", "FINAL LT", "FINAL HT", "TCBER", "3TBER", "BURNIN", "SWITCH"
    ]
    ordered = [n for p in priority for n in frames if n.strip().lower() == p.lower()] \
              + [n for n in frames if n not in priority]
    for name in ordered:
        df = frames[name]
        if df is None or df.empty or "COMPONENTID" not in df.columns:
            continue
        wo_col = _pick_wo_column(df)
        if not wo_col:
            continue
        comp_series = df["COMPONENTID"].astype(str).map(str.strip)
        for comp, wo in zip(comp_series, df[wo_col]):
            if comp in targets and not wo_map[comp]:
                w = _norm_wo_value(wo)
                if w:
                    wo_map[comp] = w
        if all(wo_map[c] for c in targets):
            break
    return wo_map

def fetch_wos_for_device(conn, device_id: str, log_fn=None) -> list[str]:
    """
    Given a Device ID, return all distinct WO values from MASTER_WO_TABLE.
    """
    cursor = conn.cursor()
    sql = (
        f"SELECT DISTINCT LTRIM(RTRIM(WO)) AS WO "
        f"FROM {MASTER_WO_TABLE} "
        f"WHERE LTRIM(RTRIM(Device)) = ?"
    )
    try:
        cursor.execute(sql, (device_id,))
        rows = cursor.fetchall()
        wos = [str(r[0]).strip() for r in rows if r[0] is not None and str(r[0]).strip()]
        if log_fn:
            log_fn(f"  • Device {device_id}: fetched {len(wos)} WO(s)")
        return sorted(set(wos))
    except Exception as e:
        if log_fn:
            log_fn(f"  • Device→WO lookup error: {e}")
        raise

# ───────────────────────────────────────────────────────────────
# Summary building

def summarize_final_ats(df: pd.DataFrame, ddmi_gate_sid: float | None) -> str:
    """
    Final Test PASS rule with ATS treated as RT:

    - Apply DDMI gate if available.
    - Map CHNumber to lane + stage (RT / LT / HT), with:
          1_ATS     → RT
          1_ATS_LT  → LT
          1_ATS_HT  → HT
    - If we have RT/LT/HT for all 8 lanes and all rows PASS → 'PASS'.
    - Otherwise:
        * Non-PASS rows → CHx code
        * Missing temps → 'CHx missing RT/LT/HT'
    """
    if df is None or df.empty:
        return ""

    # Normalize & copy
    df = normalize_test_df(df).copy()

    # Apply DDMI gate — fall back to unfiltered if gate removes everything
    if ddmi_gate_sid is not None and "_SID_NUM_" in df.columns:
        gated = df[pd.to_numeric(df["_SID_NUM_"], errors="coerce") >= ddmi_gate_sid]
        if not gated.empty:
            df = gated

    # Lane index (1..8)
    df["_CH_IDX_"] = pd.to_numeric(df.get("_CH_IDX_"), errors="coerce")
    df = df[df["_CH_IDX_"].between(1, 8)]
    if df.empty:
        return ""

    # Decide stage per row (ATS → RT for Final Test)
    df["_STAGE_"] = [
        _stage_from_row_for_station("Final Test",
                                    row.get("CHNumber"),
                                    row.get("_STAGE_"))
        for _, row in df.iterrows()
    ]
    df["_STAGE_"] = df["_STAGE_"].astype(str).str.upper()
    df = df[df["_STAGE_"].isin(["RT", "LT", "HT"])]
    if df.empty:
        return ""

    required = {"RT", "LT", "HT"}

    by_lane_stages = (
        df.groupby("_CH_IDX_")["_STAGE_"]
          .apply(lambda s: set(x for x in s if x))
          .to_dict()
    )

    status = df["CH_Pass_Fail"].astype(str).str.upper().str.strip()

    # If all lanes have RT/LT/HT and every row is PASS → PASS
    full_coverage = True
    for lane in range(1, 9):
        have = by_lane_stages.get(lane, set())
        if not required.issubset(have):
            full_coverage = False
            break

    if full_coverage and (status == "PASS").all():
        return "PASS"

    # Otherwise build failure / missing messages
    msgs = []
    for lane, grp in df.groupby("_CH_IDX_"):
        lane_tag = f"CH{int(lane)}"

        # Non-PASS failure codes
        st = grp["CH_Pass_Fail"].astype(str).str.upper().str.strip()
        codes = grp.loc[st != "PASS", "FailureCodeID"].astype(str).str.strip()
        codes = [
            c for c in codes
            if c not in ("", "NAN", "NONE", "NULL", "nan", "NaN", "null")
        ]
        if codes:
            uniq, seen = [], set()
            for c in codes:
                if c not in seen:
                    seen.add(c)
                    uniq.append(c)
            msgs.append(f"{lane_tag} " + "/".join(uniq))

        # Missing temps (after ATS→RT mapping)
        missing = required - by_lane_stages.get(lane, set())
        if missing:
            msgs.append(f"{lane_tag} missing " + "/".join(sorted(missing)))

    return ", ".join(msgs) if msgs else ""


def summarize_switch(df: pd.DataFrame) -> str:
    """Switch Test: PASS if all CH_Pass_Fail rows are PASS, else report fail codes.
    No lane-count requirement — Switch Test rows may not follow the 8-channel pattern."""
    if df is None or df.empty:
        return ""
    df = normalize_test_df(df)
    status = df["CH_Pass_Fail"].astype(str).str.upper().str.strip()
    tested = status[status != ""]
    if tested.empty:
        return ""
    if (tested == "PASS").all():
        return "PASS"
    fail_df = df[status.isin(["", "PASS"]) == False].copy()
    msgs = []
    for _, r in fail_df.iterrows():
        code = str(r.get("FailureCodeID", "")).strip()
        if code.lower() in ("", "nan", "none", "null"):
            code = str(r["CH_Pass_Fail"]).strip()
        ch = r.get("_CH_IDX_")
        prefix = f"CH{int(ch)} " if ch and not pd.isna(ch) else ""
        if code:
            msgs.append(f"{prefix}{code}")
    return ", ".join(msgs) if msgs else "FAIL"


def build_summary(type_to_frames: dict, components: list, wo_map: dict[str, str]) -> pd.DataFrame:
    trx_map = type_to_frames.get("TRX", {})
    other   = {k: v for k, v in type_to_frames.items() if k != "TRX"}
    ddmi_gate_map = compute_ddmi_gate_sid_for_components(trx_map, components)
    rows = []
    for comp in components:
        comp_key = str(comp).strip()
        row = {"WO": wo_map.get(comp_key, ""), "SN": comp_key}
        gate_sid = ddmi_gate_map.get(comp_key)
        for colname, (family, subtype) in SUMMARY_COLUMNS:
            if family == "TRX":
                df = trx_map.get(subtype, pd.DataFrame())
                sub = (
                    df[df["COMPONENTID"].astype(str).str.strip() == comp_key]
                    if isinstance(df, pd.DataFrame) and not df.empty
                    else pd.DataFrame()
                )

                # Special rule for Final Test
                if subtype == "FINAL":
                    row[colname] = summarize_final_ats(sub, gate_sid)
                else:
                    row[colname] = summarize_one(sub, gate_sid, apply_gate=(subtype not in ("DDMI",)))

            elif family in {"TCBER", "3TBER"}:
                df = other.get(family, pd.DataFrame())
                sub = (
                    df[df["COMPONENTID"].astype(str).str.strip() == comp_key]
                    if isinstance(df, pd.DataFrame) and not df.empty
                    else pd.DataFrame()
                )
                row[colname] = summarize_ber(sub, gate_sid)

            elif family == "FWWRITE":
                df = other.get("FWWRITE", pd.DataFrame())
                row[colname] = summarize_fw_write(df, comp_key, wo_map.get(comp_key, ""))
            elif family == "MODEHOP":
                df = other.get("MODEHOP", pd.DataFrame())
                row[colname] = summarize_mode_hopping(df, comp_key)
            elif family == "BURNIN":
                df = other.get("BURNIN", pd.DataFrame())
                sub = (
                    df[df["COMPONENTID"].astype(str).str.strip() == comp_key]
                    if isinstance(df, pd.DataFrame) and not df.empty
                    else pd.DataFrame()
                )
                row[colname] = summarize_burnin(sub, gate_sid)

            elif family == "SWITCH":
                df = other.get("SWITCH", pd.DataFrame())
                sub = (
                    df[df["COMPONENTID"].astype(str).str.strip() == comp_key]
                    if isinstance(df, pd.DataFrame) and not df.empty
                    else pd.DataFrame()
                )
                row[colname] = summarize_switch(sub)

            else:
                df = other.get(family, pd.DataFrame())
                sub = (
                    df[df["COMPONENTID"].astype(str).str.strip() == comp_key]
                    if isinstance(df, pd.DataFrame) and not df.empty
                    else pd.DataFrame()
                )
                row[colname] = summarize_one(sub, gate_sid, apply_gate=True)

        rows.append(row)

    cols = ["WO", "SN"] + [c for c, _ in SUMMARY_COLUMNS]
    out = pd.DataFrame(rows, columns=cols)
    return out.sort_values(by=["WO", "SN"], na_position="last").reset_index(drop=True)


def compute_fpy(summary_df: pd.DataFrame):
    """Return (first_pass_count, total_count).
    A device is first-pass if every station cell is empty or 'PASS'."""
    if summary_df is None or summary_df.empty:
        return 0, 0
    test_cols = [c for c in summary_df.columns if c not in ("WO", "SN")]
    if not test_cols:
        return 0, 0
    total = len(summary_df)
    vals = summary_df[test_cols].fillna("").astype(str).apply(
        lambda col: col.str.strip().str.upper()
    )
    fp_mask = vals.apply(lambda row: all(v in ("", "PASS") for v in row), axis=1)
    return int(fp_mask.sum()), total

# ───────────────────────────────────────────────────────────────
# Excel writer (compare formatting)

def render_ddmi_compare_sheet(wb, sheet_name, ddmi_blocks, components, active_metrics, active_sheets):
    from openpyxl.styles import Alignment, Font, Border, Side
    from openpyxl.utils import get_column_letter

    if sheet_name in wb.sheetnames:
        wb.remove(wb[sheet_name])
    ws = wb.create_sheet(sheet_name)

    by_sn_metric = {}
    sn_to_wo = {}
    for wo, sn, metric, df in ddmi_blocks:
        if metric in active_metrics:
            by_sn_metric[(sn, metric)] = df
            sn_to_wo[sn] = wo

    group_headers = active_sheets + ["std across"]
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    r = 1
    for sn in components:
        snk = str(sn).strip()
        metrics_for_sn = [m for m in active_metrics if (snk, m) in by_sn_metric]
        if not metrics_for_sn:
            continue

        wo = sn_to_wo.get(snk, "")

        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3 + len(group_headers)*len(metrics_for_sn))
        ws.cell(r, 1, f"WO: {wo}   SN: {snk}").font = Font(bold=True, size=12)
        r += 1

        # Header row 1
        c = 1
        for hdr in ["wo", "sn", "Sheet"]:
            ws.cell(r, c, hdr).font = Font(bold=True)
            ws.cell(r, c).alignment = Alignment(horizontal="center")
            c += 1
        for _metric in metrics_for_sn:
            for hdr in group_headers:
                cell = ws.cell(r, c, hdr)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center")
                cell.border = border
                c += 1
        r += 1

        # Header row 2
        c = 1
        ws.cell(r, c, wo).alignment = Alignment(horizontal="center"); c += 1
        ws.cell(r, c, snk).alignment = Alignment(horizontal="center"); c += 1
        ws.cell(r, c, "Metric").font = Font(italic=True)
        ws.cell(r, c).alignment = Alignment(horizontal="center"); c += 1
        for metric in metrics_for_sn:
            for _ in group_headers:
                cell = ws.cell(r, c, metric)
                cell.font = Font(italic=True)
                cell.alignment = Alignment(horizontal="center")
                cell.border = border
                c += 1
        r += 1

        # Body CH1..CH8
        for ch in range(1, 9):
            c = 1
            ws.cell(r, c, wo).alignment = Alignment(horizontal="center"); c += 1
            ws.cell(r, c, snk).alignment = Alignment(horizontal="center"); c += 1
            ws.cell(r, c, f"CH{ch}").font = Font(bold=True)
            ws.cell(r, c).alignment = Alignment(horizontal="center"); c += 1

            for metric in metrics_for_sn:
                block = by_sn_metric[(snk, metric)]
                row = block[block["CH"] == f"CH{ch}"]
                if row.empty:
                    vals = [None] * (len(active_sheets) + 1)
                else:
                    rr = row.iloc[0]
                    vals = [rr.get(s) for s in active_sheets] + [rr.get("StdAcross")]
                for v in vals:
                    cell = ws.cell(r, c, v)
                    cell.alignment = Alignment(horizontal="center")
                    cell.border = border
                    c += 1
            r += 1

        for col in range(1, 4 + len(group_headers)*len(metrics_for_sn)):
            ws.column_dimensions[get_column_letter(col)].width = 14
        r += 1

# ───────────────────────────────────────────────────────────────
# ───────────────────────────────────────────────────────────────
# Models / UI bits

class DataFrameModel(QAbstractTableModel):
    """
    Generic DataFrame model with optional per-SN marker colouring.
    marker_colors_by_sn: dict[str, QColor]
    """
    def __init__(self, df=pd.DataFrame(), parent=None, marker_colors_by_sn=None):
        super().__init__(parent)
        self._df = df
        self._marker_colors_by_sn = marker_colors_by_sn or {}

    def setDataFrame(self, df: pd.DataFrame):
        self.beginResetModel()
        self._df = df.copy()
        self.endResetModel()

    def setMarkers(self, marker_map: dict | None):
        """Update mapping from SN -> QColor used for row background."""
        self.beginResetModel()
        self._marker_colors_by_sn = marker_map or {}
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._df)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._df.columns)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None

        value = self._df.iat[index.row(), index.column()]

        # Text
        if role in (Qt.DisplayRole, Qt.EditRole):
            return "" if pd.isna(value) else str(value)

        # Background color (Markers first, then PASS/FAIL)
        if role == Qt.BackgroundRole:
            # Marker-based row highlight
            if self._marker_colors_by_sn and "SN" in self._df.columns:
                try:
                    sn_col_idx = list(self._df.columns).index("SN")
                    sn_val = self._df.iat[index.row(), sn_col_idx]
                    sn_key = "" if pd.isna(sn_val) else str(sn_val).strip()
                    if sn_key in self._marker_colors_by_sn:
                        return QBrush(self._marker_colors_by_sn[sn_key])
                except Exception:
                    pass

            # Normal PASS / FAIL colouring
            try:
                col_name = str(self._df.columns[index.column()])
            except Exception:
                col_name = ""

            if pd.isna(value):
                return None

            if isinstance(value, str):
                v_str = value.strip()
            else:
                v_str = str(value).strip()

            v_upper = v_str.upper()

            if v_upper == "PASS":
                return QBrush(QColor("#C6EFCE"))  # green

            if v_str not in ("", " ") and col_name not in ("WO", "SN"):
                return QBrush(QColor("#FFC7CE"))  # red

        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return str(self._df.columns[section])
        return str(section + 1)


class ElideDelegate(QStyledItemDelegate):
    def __init__(self, row_h=26, parent=None):
        super().__init__(parent)
        self._row_h = row_h

    def paint(self, painter, option, index):
        opt = QStyleOptionViewItem(option)
        opt.textElideMode = Qt.ElideRight
        super().paint(painter, opt, index)

    def sizeHint(self, option, index):
        s = super().sizeHint(option, index)
        return QSize(s.width(), self._row_h)


class DetailDataFrameModel(DataFrameModel):
    """
    Detail model used only in the raw-data tab of the popup.
    Colour logic:
      - Only CH_Pass_Fail / Pass/Fail columns are coloured.
      - PASS  -> green
      - any non-empty non-PASS -> red
    """
    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None

        value = self._df.iat[index.row(), index.column()]

        if role in (Qt.DisplayRole, Qt.EditRole):
            return "" if pd.isna(value) else str(value)

        if role == Qt.BackgroundRole:
            try:
                col_name = str(self._df.columns[index.column()])
            except Exception:
                col_name = ""

            if col_name in ("CH_Pass_Fail", "Pass/Fail"):
                v_str = "" if pd.isna(value) else str(value).strip().upper()
                if v_str == "PASS":
                    return QBrush(QColor("#C6EFCE"))
                elif v_str not in ("", " "):
                    return QBrush(QColor("#FFC7CE"))
            return None

        return None


# --- NEW: pivot-table model for station popup ----------------

class StationPivotModel(QAbstractTableModel):
    """
    Pivot view for a station:
      columns: Metric, MAX, MIN, CH1..CH8
      rows: metrics, filled metrics first, then empty ones.
    CH1..CH8 cells are coloured red if they are outside
    the Min/Max limits from STATION_SPECS for this station+metric.
    """
    def __init__(self, df: pd.DataFrame, station_name: str, raw_df: pd.DataFrame | None = None, parent=None):
        super().__init__(parent)
        self._df = df.copy()
        self._station_name = station_name
        # raw_df is the original raw rows used to build this pivot (may be filtered to a TestNumber).
        # Use it to detect if ALL channels passed for the device at that test run; if so, skip colouring.
        self._raw_df = raw_df.copy() if raw_df is not None and not raw_df.empty else None
        self._all_channels_passed = False
        try:
            if self._raw_df is not None:
                d = normalize_test_df(self._raw_df.copy())
                # Need a channel index for grouping
                if "_CH_IDX_" not in d.columns and "CHNumber" in d.columns:
                    d["_CH_IDX_"] = d["CHNumber"].apply(lambda x: parse_lane_stage(x)[0])
                # Consider channels 1..8 that appear in data; check latest per channel
                ch_pass = []
                if "_CH_IDX_" in d.columns:
                    for ch in range(1, 9):
                        mask = d["_CH_IDX_"] == ch
                        if not mask.any():
                            continue
                        sub = d[mask]
                        # use numeric SID if present to pick latest
                        if "_SID_NUM_" in sub.columns:
                            sub = sub.sort_values("_SID_NUM_")
                        # pick last row's pass value
                        v = None
                        if "CH_Pass_Fail" in sub.columns:
                            v = str(sub.iloc[-1].get("CH_Pass_Fail", "")).strip().upper()
                        elif "Pass/Fail" in sub.columns:
                            v = str(sub.iloc[-1].get("Pass/Fail", "")).strip().upper()
                        else:
                            v = None
                        if v is None or v == "":
                            # treat missing as not-passed
                            ch_pass.append(False)
                        else:
                            ch_pass.append(v == "PASS")
                # If we have at least one channel and all are True, set flag
                if ch_pass and all(ch_pass):
                    self._all_channels_passed = True
        except Exception:
            self._all_channels_passed = False

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._df)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._df.columns)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None

        value = self._df.iat[index.row(), index.column()]

        if role in (Qt.DisplayRole, Qt.EditRole):
            return "" if pd.isna(value) else str(value)

        if role == Qt.BackgroundRole:
            try:
                col_name = str(self._df.columns[index.column()])
            except Exception:
                col_name = ""

            # Only colour CH1..CH8 based on spec limits
            if col_name.startswith("CH") and "Metric" in self._df.columns:
                # If all channels passed for this device/test run, skip colouring
                if self._all_channels_passed:
                    return None

                metric_name = str(self._df.at[index.row(), "Metric"])
                limits = STATION_SPECS.get(self._station_name, {}).get(metric_name)
                if limits and not pd.isna(value):
                    lo, hi = limits
                    try:
                        v = float(value)
                    except Exception:
                        return None
                    if (lo is not None and v < lo) or (hi is not None and v > hi):
                        return QBrush(QColor("#FFC7CE"))
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return str(self._df.columns[section])
        return str(section + 1)


def _build_station_pivot(sub: pd.DataFrame, station_name: str | None = None) -> pd.DataFrame:
    """
    Build pivot table:
      Metric | MIN | MAX | CH1..CH8

    - MIN = LSL from STATION_SPECS (if found)
    - MAX = USL from STATION_SPECS (if found)
    - CH1..CH8 = latest measured value per channel
    - Metrics with any real data appear first, then all-NaN ones.
    """
    if sub is None or sub.empty:
        return pd.DataFrame()

    df = sub.copy()

    # Ensure lane index
    if "_CH_IDX_" not in df.columns:
        if "CHNumber" in df.columns:
            df["_CH_IDX_"] = df["CHNumber"].apply(lambda x: parse_lane_stage(x)[0])
        else:
            df["_CH_IDX_"] = np.nan

    skip_cols = {
        "COMPONENTID",
        "SID",
        "CHNumber",
        "CH_Pass_Fail",
        "Pass/Fail",
        "FailureCodeID",
        "_SID_NUM_",
        "_CH_IDX_",
        "_STAGE_",
    }

    metric_cols = []
    for col in df.columns:
        if col in skip_cols:
            continue
        # treat any numeric-ish column as a metric; ordering will handle all-NaN
        _ = pd.to_numeric(df[col], errors="coerce")
        metric_cols.append(col)

    filled_rows = []
    empty_rows = []

    for metric in metric_cols:
        # Latest value per CH1..CH8
        ch_vals = []
        for ch in range(1, 9):
            mask = df["_CH_IDX_"] == ch
            s = pd.to_numeric(df.loc[mask, metric], errors="coerce").dropna()
            v = float(s.iloc[-1]) if not s.empty else np.nan
            ch_vals.append(v)

        arr = np.array(ch_vals, dtype=float)

        # Default: NaNs for spec
        min_v = np.nan
        max_v = np.nan

        # Lookup spec for this station + metric
        limits = None
        if station_name:
            limits = STATION_SPECS.get(station_name, {}).get(str(metric))

        if limits is not None:
            lo, hi = limits
            # Use spec limits (LSL/USL) directly
            min_v = lo if lo is not None else np.nan
            max_v = hi if hi is not None else np.nan
        else:
            # Fallback: use measured extremes if no spec defined
            if not np.all(np.isnan(arr)):
                min_v = float(np.nanmin(arr))
                max_v = float(np.nanmax(arr))

        row = [metric, min_v, max_v] + ch_vals

        if np.all(np.isnan(arr)):
            empty_rows.append(row)
        else:
            filled_rows.append(row)

    rows = filled_rows + empty_rows
    cols = ["Metric", "MIN", "MAX"] + [f"CH{i}" for i in range(1, 9)]
    return pd.DataFrame(rows, columns=cols)



class StationDetailDialog(QDialog):
    """
    Popup that shows:
      - Pivot view (metrics vs channels) with SID selector
      - Raw rows
    for one SN + station.
    """
    def __init__(self, sn, station_name, df, raw_all_df=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"{station_name} – SN {sn}")
        self.resize(1100, 650)
        self._station_name = station_name
        self._reduced_df = df
        self._raw_all_df = raw_all_df if raw_all_df is not None and not raw_all_df.empty else df

        lay = QVBoxLayout(self)

        info = QLabel(f"SN: {sn}   Station: {station_name}   Rows: {len(df)}")
        lay.addWidget(info)

        # Test-run selector row
        sid_row = QHBoxLayout()
        sid_row.addWidget(QLabel("Test run:"))
        self._sid_combo = QComboBox()
        self._sid_combo.setMinimumWidth(220)
        self._sid_items = _build_test_run_list(self._raw_all_df)
        for label, _rank in self._sid_items:
            self._sid_combo.addItem(label)
        sid_row.addWidget(self._sid_combo)
        sid_row.addStretch(1)
        lay.addLayout(sid_row)

        tabs = QTabWidget()
        lay.addWidget(tabs, 1)

        # Tab 1: Pivot view
        pivot_df = _build_station_pivot(df, station_name)
        self.pivot_view = QTableView()
        if not pivot_df.empty:
            # pass the raw rows `df` so the model can decide to skip colouring
            self.pivot_model = StationPivotModel(pivot_df, station_name, raw_df=df)
            self.pivot_view.setModel(self.pivot_model)
            compact_table(self.pivot_view, row_h=26, min_col_w=110, first_col_w=160)
        else:
            self.pivot_model = None
        w1 = QWidget()
        v1 = QVBoxLayout(w1)
        v1.addWidget(self.pivot_view)
        tabs.addTab(w1, "Pivot view")

        # Tab 2: Raw data
        self.raw_model = DetailDataFrameModel(df)
        self.raw_view = QTableView()
        self.raw_view.setModel(self.raw_model)
        compact_table(self.raw_view, row_h=26, min_col_w=110, first_col_w=140)
        w2 = QWidget()
        v2 = QVBoxLayout(w2)
        v2.addWidget(self.raw_view)
        tabs.addTab(w2, "Raw data")

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_close = QPushButton("Close")
        btn_close.clicked.connect(self.accept)
        btn_row.addWidget(btn_close)
        lay.addLayout(btn_row)

        # Connect combo after UI is built
        self._sid_combo.currentIndexChanged.connect(self._on_sid_changed)

    def _on_sid_changed(self, idx):
        if idx < 0 or idx >= len(self._sid_items):
            return
        _label, tn_val = self._sid_items[idx]
        if tn_val is None:
            filtered = self._reduced_df
        else:
            tn_col = _find_testnumber_col(self._raw_all_df)
            if tn_col:
                filtered = self._raw_all_df[
                    self._raw_all_df[tn_col].astype(str).str.strip() == tn_val
                ].copy()
            else:
                filtered = self._reduced_df
        pivot_df = _build_station_pivot(filtered, self._station_name)
        if not pivot_df.empty:
            self.pivot_model = StationPivotModel(pivot_df, self._station_name, raw_df=filtered)
        else:
            self.pivot_model = StationPivotModel(pd.DataFrame(columns=["Metric","MIN","MAX"]+[f"CH{i}" for i in range(1,9)]), self._station_name, raw_df=filtered)
        self.pivot_view.setModel(self.pivot_model)
        compact_table(self.pivot_view, row_h=26, min_col_w=110, first_col_w=160)
class MultiStageDetailDialog(QDialog):
    """
    Popup with one tab per temperature stage (RT / LT / HT) for a single SN + station.

    stage_to_df: dict[label -> DataFrame]
    """
    def __init__(self, sn, station_name, stage_to_df, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"{station_name} – SN {sn}")
        self.resize(1100, 650)

        layout = QVBoxLayout(self)
        tabs = QTabWidget()
        layout.addWidget(tabs)

        # One tab per stage label (e.g. "RT", "LT", "HT", or "RT (ATS)")
        for label, df in stage_to_df.items():
            page = QWidget()
            v = QVBoxLayout(page)

            info = QLabel(f"Rows: {len(df)}")
            v.addWidget(info)

            model = DetailDataFrameModel(df)
            view = QTableView()
            view.setModel(model)
            compact_table(view, row_h=26, min_col_w=110, first_col_w=140)
            v.addWidget(view)

            tabs.addTab(page, label)

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_close = QPushButton("Close")
        btn_close.clicked.connect(self.accept)
        btn_row.addWidget(btn_close)
        layout.addLayout(btn_row)

def _split_three_temps_for_station(station_name: str,
                                   df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """
    Take raw rows for one SN+station (all temps) and return:
        {"RT": df_rt, "LT": df_lt, "HT": df_ht}

    For 3TBER / TCBER: use _STAGE_ parsed from CHNumber (RT/LT/HT).
    For Final Test: treat ATS as RT (1_ATS → RT, 1_ATS_LT → LT, 1_ATS_HT → HT).
    """
    if df is None or df.empty:
        return {"RT": pd.DataFrame(), "LT": pd.DataFrame(), "HT": pd.DataFrame()}

    df = normalize_test_df(df).copy()

    # Ensure lane index
    if "_CH_IDX_" not in df.columns:
        if "CHNumber" in df.columns:
            df["_CH_IDX_"] = df["CHNumber"].apply(lambda x: parse_lane_stage(x)[0])
        else:
            df["_CH_IDX_"] = np.nan
    df["_CH_IDX_"] = pd.to_numeric(df["_CH_IDX_"], errors="coerce")
    df = df[df["_CH_IDX_"].between(1, 8)]
    if df.empty:
        return {"RT": pd.DataFrame(), "LT": pd.DataFrame(), "HT": pd.DataFrame()}

    # Decide stage
    if station_name in ("3TBER", "TCBER"):
        df = _ensure_stage_col(df)
        df["_STAGE_"] = df["_STAGE_"].astype(str).str.upper()
    elif station_name == "Final Test":
        df["_STAGE_"] = [
            _stage_from_row_for_station("Final Test",
                                        row.get("CHNumber"),
                                        row.get("_STAGE_"))
            for _, row in df.iterrows()
        ]
        df["_STAGE_"] = df["_STAGE_"].astype(str).str.upper()
    else:
        df = _ensure_stage_col(df)
        df["_STAGE_"] = df["_STAGE_"].astype(str).str.upper()

    out = {}
    for tag in ("RT", "LT", "HT"):
        out[tag] = df[df["_STAGE_"] == tag].copy()
    return out

class ThreeTempPivotDialog(QDialog):
    """
    Popup for 3TBER / TCBER / Final Test:

      Tabs:
        - RT: pivot (Metric, MAX, MIN, CH1..CH8) with SID selector
        - LT: pivot
        - HT: pivot
        - Raw: all temps combined, raw rows

    Pivot uses StationPivotModel + STATION_SPECS for colouring.
    """
    def __init__(self, sn: str, station_name: str,
                 raw_df: pd.DataFrame, raw_all_df=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"{station_name} – SN {sn}")
        self.resize(1150, 680)
        self._station_name = station_name
        self._reduced_df = raw_df
        self._raw_all_df = raw_all_df if raw_all_df is not None and not raw_all_df.empty else raw_df

        main = QVBoxLayout(self)

        lbl = QLabel(f"SN: {sn}    Station: {station_name}    Rows after gate: {len(raw_df)}")
        main.addWidget(lbl)

        # Test-run selector row
        sid_row = QHBoxLayout()
        sid_row.addWidget(QLabel("Test run:"))
        self._sid_combo = QComboBox()
        self._sid_combo.setMinimumWidth(220)
        self._sid_items = _build_test_run_list(self._raw_all_df)
        for label, _rank in self._sid_items:
            self._sid_combo.addItem(label)
        sid_row.addWidget(self._sid_combo)
        sid_row.addStretch(1)
        main.addLayout(sid_row)

        self._tabs = QTabWidget()
        main.addWidget(self._tabs, 1)

        # Build initial pivot tabs from reduced data
        self._stage_views = {}
        self._stage_info_labels = {}
        self._build_pivot_tabs(raw_df)

        # Raw data tab (all temps together)
        raw_model = DetailDataFrameModel(raw_df)
        raw_view = QTableView()
        raw_view.setModel(raw_model)
        compact_table(raw_view, row_h=26, min_col_w=110, first_col_w=140)

        raw_page = QWidget()
        rv = QVBoxLayout(raw_page)
        rv.addWidget(raw_view)
        self._tabs.addTab(raw_page, "Raw")

        # Close button
        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_close = QPushButton("Close")
        btn_close.clicked.connect(self.accept)
        btn_row.addWidget(btn_close)
        main.addLayout(btn_row)

        # Connect combo after UI is built
        self._sid_combo.currentIndexChanged.connect(self._on_sid_changed)

    def _build_pivot_tabs(self, source_df):
        """Build/rebuild RT/LT/HT pivot tabs from source_df."""
        # Remove old stage tabs (keep Raw tab at end)
        for stage in list(self._stage_views.keys()):
            for i in range(self._tabs.count()):
                if self._tabs.tabText(i) == stage:
                    self._tabs.removeTab(i)
                    break
        self._stage_views.clear()
        self._stage_info_labels.clear()

        stage_to_df = _split_three_temps_for_station(self._station_name, source_df)

        insert_idx = 0
        for stage in ("RT", "LT", "HT"):
            stage_df = stage_to_df.get(stage)
            if stage_df is None or stage_df.empty:
                continue

            pivot_df = _build_station_pivot(stage_df, self._station_name)
            if pivot_df.empty:
                continue
            view = QTableView()
            # pass stage raw rows so pivot can skip colouring if all channels passed
            model = StationPivotModel(pivot_df, self._station_name, raw_df=stage_df)
            view.setModel(model)
            compact_table(view, row_h=26, min_col_w=110, first_col_w=170)

            page = QWidget()
            v = QVBoxLayout(page)
            info = QLabel(f"{stage}: {len(stage_df)} raw rows \u2192 {len(pivot_df)} metrics")
            v.addWidget(info)
            v.addWidget(view)

            self._tabs.insertTab(insert_idx, page, stage)
            self._stage_views[stage] = view
            self._stage_info_labels[stage] = info
            insert_idx += 1

    def _on_sid_changed(self, idx):
        if idx < 0 or idx >= len(self._sid_items):
            return
        _label, tn_val = self._sid_items[idx]
        if tn_val is None:
            source = self._reduced_df
        else:
            tn_col = _find_testnumber_col(self._raw_all_df)
            if tn_col:
                source = self._raw_all_df[
                    self._raw_all_df[tn_col].astype(str).str.strip() == tn_val
                ].copy()
            else:
                source = self._reduced_df
        self._build_pivot_tabs(source)

def copy_table_selection_to_clipboard(tv: QTableView):
    sm = tv.selectionModel()
    model = tv.model()
    if sm is None or model is None:
        return

    idxs = sm.selectedIndexes()
    if not idxs:
        return

    # Sort selection
    idxs = sorted(idxs, key=lambda i: (i.row(), i.column()))

    rows = [i.row() for i in idxs]
    cols = [i.column() for i in idxs]
    r0, r1 = min(rows), max(rows)
    c0, c1 = min(cols), max(cols)

    idx_map = {(i.row(), i.column()): i for i in idxs}

    lines = []
    for r in range(r0, r1 + 1):
        row_vals = []
        for c in range(c0, c1 + 1):
            ix = idx_map.get((r, c))
            if ix is None:
                row_vals.append("")
            else:
                v = model.data(ix, Qt.DisplayRole)
                row_vals.append("" if v is None else str(v))
        lines.append("\t".join(row_vals))

    QApplication.clipboard().setText("\n".join(lines))

def compact_table(tv: QTableView, row_h=26, min_col_w=110, first_col_w=140):
    tv.setAlternatingRowColors(True)
    tv.setWordWrap(False)
    tv.setTextElideMode(Qt.ElideRight)
    tv.setItemDelegate(ElideDelegate(row_h=row_h, parent=tv))
    vh = tv.verticalHeader()
    vh.setVisible(False)
    vh.setDefaultSectionSize(row_h)
    hh = tv.horizontalHeader()
    hh.setStretchLastSection(False)
    hh.setMinimumSectionSize(60)
    hh.setDefaultSectionSize(min_col_w)
    hh.setSectionResizeMode(QHeaderView.Interactive)
    if tv.model():
        for c in range(tv.model().columnCount()):
            tv.setColumnWidth(c, first_col_w if c == 0 else min_col_w)
        # Allow multi-cell selection
    tv.setSelectionMode(QAbstractItemView.ExtendedSelection)

    # Install Excel-style Ctrl+C copy (only once per table)
    if not getattr(tv, "_copy_action_installed", False):
        act_copy = QAction("Copy", tv)
        act_copy.setShortcut(QKeySequence.Copy)
        # Important: works even if a cell editor has focus
        act_copy.setShortcutContext(Qt.WidgetWithChildrenShortcut)
        act_copy.triggered.connect(lambda _=False, t=tv: copy_table_selection_to_clipboard(t))
        tv.addAction(act_copy)
        tv._copy_action_installed = True


def build_app_qss(font_px=13, row_h=26, header_h=30, cell_pad_h=3, cell_pad_w=8, prog_h=12):
    return f"""
* {{
  font-family: 'Segoe UI','Inter','Arial';
  font-size: {font_px}px;
}}
QWidget {{
  background: #FFFFFF;
  color: #111827;
}}
QGroupBox {{
  border: 1px solid #E2E6EC;
  border-radius: 6px;
  margin-top: 8px;
  padding: 6px 8px 8px 8px;
}}
QGroupBox::title {{
  subcontrol-origin: margin;
  left: 10px;
  padding: 0 4px;
  color: #6B7280;
}}
QHeaderView::section {{
  background: #F2F4F7;
  color: #111827;
  padding: {cell_pad_h}px {cell_pad_w}px;
  border: 1px solid #E2E6EC;
  font-weight: 600;
  min-height: {header_h}px;
}}
QTableView {{
  background: #FFFFFF;
  gridline-color: #E2E6EC;
  alternate-background-color: #F9FAFB;
}}
QTableView::item {{
  padding: {cell_pad_h}px {cell_pad_w}px;
}}
QTableView::item:selected {{
  background: #E7F0FF;
  color: #111827;
}}
QScrollBar:vertical, QScrollBar:horizontal {{
  background: transparent;
  border: none;
  min-width: 12px;
  min-height: 12px;
}}
QScrollBar::handle:vertical, QScrollBar::handle:horizontal {{
  background: #D9DEE7;
  border-radius: 6px;
}}
QScrollBar::handle:hover {{
  background: #C7D3EA;
}}
QPushButton {{
  background: #FFFFFF;
  border: 1px solid #D9DEE7;
  border-radius: 6px;
  padding: 7px 14px;
}}
QPushButton:hover {{
  background: #F8FAFF;
  border-color: #C7D3EA;
}}
QPushButton:pressed {{
  background: #EEF2FF;
}}
QComboBox, QSpinBox, QLineEdit {{
  background: #FFFFFF;
  border: 1px solid #D9DEE7;
  border-radius: 6px;
  padding: 4px 8px;
}}
QTextEdit {{
  background: #FFFFFF;
  border: 1px solid #D9DEE7;
  border-radius: 6px;
  padding: 6px 8px;
}}
QProgressBar {{
  background: #EFF3FA;
  border: 1px solid #E0E5EF;
  border-radius: 6px;
  text-align: center;
  height: {prog_h}px;
}}
QProgressBar::chunk {{
  background-color: #5AA8FF;
  border-radius: 6px;
}}
"""

# ───────────────────────────────────────────────────────────────
# Worker (SQL path shared by both tabs)

class QueryWorker(QThread):
    progress = pyqtSignal(str)
    progressPct = pyqtSignal(int)
    done = pyqtSignal(dict, list, dict, pd.DataFrame, pd.DataFrame, list)
    error = pyqtSignal(str)

    def __init__(self, mode_is_sn: bool, sn_list: list[str], wo_lines: list[str],
                 build_compare: bool, restrict_device_id: str | None = None, parent=None):
        super().__init__(parent)
        self.mode_is_sn = mode_is_sn
        self.sn_list = sn_list
        self.wo_lines = wo_lines
        self.build_compare = build_compare  # control whether to build Compare tables
        # For Search SN only: restrict WO mapping to a device (DR8+/FR4) so SN doesn't map to "random" WO
        self.restrict_device_id = (restrict_device_id or "").strip() or None

    def run(self):
        try:
            step = 0
            def bump(msg, add=8):
                nonlocal step
                step = min(100, step + add)
                self.progress.emit(msg)
                self.progressPct.emit(step)

            bump("Connecting to SQL…", 8)
            conn = pyodbc.connect(DB_CONN)

            # Resolve components
            if self.mode_is_sn:
                comps = [str(x).strip() for x in self.sn_list if str(x).strip()]
                if not comps:
                    raise RuntimeError("Enter at least one COMPONENTID.")
            else:
                raw_wos = [p.strip() for line in self.wo_lines for p in line.split(",") if p.strip()]
                if not raw_wos:
                    raise RuntimeError("Enter one or more Work Orders.")
                comps = fetch_components_by_wos(conn, raw_wos, log_fn=self.progress.emit)
                if not comps:
                    raise RuntimeError("No COMPONENTIDs found for those WO(s).")

            bump("IDs ready.", 4)

            # Fetch all
            raw_type_to_frames: dict = {}

            self.progress.emit("Querying TRX (bulk)…")
            trx_df_raw = fetch_table_for_components_bulk(conn, TRX_TABLE, comps, log_fn=self.progress.emit)
            trx_df = keep_latest_per_channel(trx_df_raw) if not trx_df_raw.empty else trx_df_raw
            trx_split = split_trx_types(trx_df)
            type_to_frames: dict = {"TRX": trx_split}
            raw_type_to_frames["TRX"] = split_trx_types(normalize_test_df(trx_df_raw) if not trx_df_raw.empty else trx_df_raw, reduce=False)
            bump("TRX fetched.", 8)

            # FW Writing + Mode Hopping (separate tables)
            self.progress.emit("Querying FW Writing (bulk)…")
            fw_df_raw = fetch_table_for_components_bulk(conn, FWWRITE_TABLE, comps, log_fn=self.progress.emit)
            fw_df = fw_df_raw.copy()
            if not fw_df.empty:
                fw_df = reduce_fw_latest_per_component(fw_df)
            type_to_frames["FWWRITE"] = fw_df
            raw_type_to_frames["FWWRITE"] = normalize_test_df(fw_df_raw) if not fw_df_raw.empty else fw_df_raw
            bump("FW Writing fetched.", 4)

            self.progress.emit("Querying Mode Hopping (bulk)…")
            mh_df_raw = fetch_table_for_components_bulk(conn, MODEHOP_TABLE, comps, log_fn=self.progress.emit)
            mh_df = mh_df_raw.copy()
            if not mh_df.empty:
                mh_latest_map = fetch_latest_modehop_testnumber_map(conn, comps, log_fn=self.progress.emit)
                mh_df = reduce_modehop_to_master_latest(mh_df, mh_latest_map)
            type_to_frames["MODEHOP"] = mh_df
            raw_type_to_frames["MODEHOP"] = normalize_test_df(mh_df_raw) if not mh_df_raw.empty else mh_df_raw
            bump("Mode Hopping fetched.", 4)

            for key, table in TABLES.items():
                self.progress.emit(f"Querying {key} (bulk)…")
                df_all_raw = fetch_table_for_components_bulk(conn, table, comps, log_fn=self.progress.emit)
                df_all = df_all_raw.copy()
                if not df_all.empty:
                    if key in {"TCBER", "3TBER"}:
                        df_all = reduce_ber_latest_per_lane_stage(df_all)
                    elif key == "BURNIN":
                        df_all = reduce_burnin_latest_per_lane_cycle(df_all)
                    else:
                        df_all = keep_latest_per_channel(df_all)
                type_to_frames[key] = df_all
                raw_type_to_frames[key] = normalize_test_df(df_all_raw) if not df_all_raw.empty else df_all_raw
                bump(f"{key} fetched.", 6)

            self._raw_type_to_frames = raw_type_to_frames

            self.progress.emit(f"WO lookup from {MASTER_WO_TABLE}…")

            # WO mapping:
            # - Search WO / Fetch WO: restrict to the WO list the user selected/entered.
            # - Search SN: restrict mapping to the selected Device (DR8+/FR4) so SN doesn't map to an unrelated WO.
            if self.mode_is_sn:
                if self.restrict_device_id:
                    self.progress.emit(f"WO lookup restricted to Device {self.restrict_device_id}…")
                    wo_map = fetch_wo_map(
                        conn,
                        comps,
                        log_fn=self.progress.emit,
                        restrict_wos=None,
                        restrict_device=self.restrict_device_id,
                    )
                else:
                    wo_map = fetch_wo_map(
                        conn,
                        comps,
                        log_fn=self.progress.emit,
                        restrict_wos=None,
                    )
            else:
                raw_wos = [p.strip() for line in self.wo_lines for p in line.split(",") if p.strip()]
                wo_map = fetch_wo_map(
                    conn,
                    comps,
                    log_fn=self.progress.emit,
                    restrict_wos=raw_wos,
                )

            bump("WO map ready.", 6)


            self.progress.emit("Building SUMMARY…")
            summary_df = build_summary(type_to_frames, comps, wo_map)
            bump("SUMMARY built.", 12)
            missing_wo = summary_df[summary_df["WO"].astype(str).str.strip().eq("")]
            if not missing_wo.empty:
                self.progress.emit(
                    f"  • SUMMARY shows {len(missing_wo)} device(s) with empty WO (check presence in {MASTER_WO_TABLE})."
                )

            if self.build_compare:
                self.progress.emit("Building Compare tables…")
                live_df, blocks = build_ddmi_compare_live_and_blocks(type_to_frames, comps, wo_map)
                bump("Compare built.", 20)
            else:
                live_df, blocks = pd.DataFrame(), []
                bump("Skipping Compare build for this tab.", 10)

            try:
                conn.close()
            except Exception:
                pass

            self.progressPct.emit(100)
            self.done.emit(type_to_frames, comps, wo_map, summary_df, live_df, blocks)

        except Exception as e:
            tb = traceback.format_exc(limit=2)
            self.error.emit(f"{e}\n{tb}")

# ───────────────────────────────────────────────────────────────
# Shared mini popups

class InputPopup(QDialog):
    def __init__(self, title="Paste IDs", label="Paste COMPONENTIDs, one per line:", parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel(label))
        self.txt = QTextEdit()
        self.txt.setMinimumHeight(120)
        lay.addWidget(self.txt)
        btns = QHBoxLayout()
        self.btn_ok = QPushButton("OK")
        self.btn_cancel = QPushButton("Cancel")
        btns.addStretch(1)
        btns.addWidget(self.btn_cancel)
        btns.addWidget(self.btn_ok)
        lay.addLayout(btns)
        self.btn_ok.clicked.connect(self.accept)
        self.btn_cancel.clicked.connect(self.reject)

    def lines(self):
        return [x.strip() for x in self.txt.toPlainText().splitlines() if x.strip()]

class WoPickerDialog(QDialog):
    """
    Popup for selecting WOs from a list, with search and multi-select.
    """
    def __init__(self, wos: list[str], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Work Orders")
        self.resize(420, 500)
        self._all_wos = sorted(set(wos))

        lay = QVBoxLayout(self)

        lay.addWidget(QLabel("Search WOs:"))
        self.search_edit = QLineEdit()
        lay.addWidget(self.search_edit)

        self.list = QListWidget()
        self.list.setSelectionMode(QAbstractItemView.MultiSelection)
        for w in self._all_wos:
            QListWidgetItem(w, self.list)
        lay.addWidget(self.list, 1)

        self.chk_all = QCheckBox("Use all WOs")
        self.chk_all.setChecked(True)
        lay.addWidget(self.chk_all)

        btns = QHBoxLayout()
        self.btn_ok = QPushButton("OK")
        self.btn_cancel = QPushButton("Cancel")
        btns.addStretch(1)
        btns.addWidget(self.btn_cancel)
        btns.addWidget(self.btn_ok)
        lay.addLayout(btns)

        self.search_edit.textChanged.connect(self._apply_filter)
        self.btn_ok.clicked.connect(self.accept)
        self.btn_cancel.clicked.connect(self.reject)

    def _apply_filter(self, text: str):
        txt = text.strip().lower()
        for i in range(self.list.count()):
            item = self.list.item(i)
            item.setHidden(bool(txt and txt not in item.text().lower()))

    def selected_wos(self) -> list[str]:
        if self.chk_all.isChecked():
            return self._all_wos[:]
        items = self.list.selectedItems()
        return [it.text().strip() for it in items if it.text().strip()]

    def accept(self):
        if not self.chk_all.isChecked() and not self.list.selectedItems():
            QMessageBox.warning(self, "No WOs Selected",
                                "Select at least one WO or check 'Use all WOs'.")
            return
        super().accept()

# ───────────────────────────────────────────────────────────────
# Marker dialog (name + colour)

class MarkerDialog(QDialog):
    def __init__(self, existing_names=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Create Marker")
        self._existing = set(existing_names or [])
        self._color = QColor("#FFF3CD")  # soft yellow default

        lay = QVBoxLayout(self)

        row_name = QHBoxLayout()
        row_name.addWidget(QLabel("Marker name:"))
        self.ed_name = QLineEdit()
        row_name.addWidget(self.ed_name)
        lay.addLayout(row_name)

        row_color = QHBoxLayout()
        row_color.addWidget(QLabel("Colour:"))
        self.btn_color = QPushButton("Pick…")
        self.lbl_preview = QLabel(" ")
        self.lbl_preview.setFixedSize(40, 20)
        self._update_preview()
        row_color.addWidget(self.btn_color)
        row_color.addWidget(self.lbl_preview)
        row_color.addStretch(1)
        lay.addLayout(row_color)

        btns = QHBoxLayout()
        self.btn_ok = QPushButton("OK")
        self.btn_cancel = QPushButton("Cancel")
        btns.addStretch(1)
        btns.addWidget(self.btn_cancel)
        btns.addWidget(self.btn_ok)
        lay.addLayout(btns)

        self.btn_color.clicked.connect(self._pick_color)
        self.btn_ok.clicked.connect(self._on_ok)
        self.btn_cancel.clicked.connect(self.reject)

    def _pick_color(self):
        col = QColorDialog.getColor(self._color, self, "Choose marker colour")
        if col.isValid():
            self._color = col
            self._update_preview()

    def _update_preview(self):
        self.lbl_preview.setStyleSheet(
            f"background-color: {self._color.name()}; "
            "border: 1px solid #D1D5DB; border-radius: 3px;"
        )

    def _on_ok(self):
        name = self.ed_name.text().strip()
        if not name:
            QMessageBox.warning(self, "Name required", "Please enter a marker name.")
            return
        if name in self._existing:
            QMessageBox.warning(self, "Duplicate name", "A marker with this name already exists.")
            return
        self.accept()

    def marker_name(self) -> str:
        return self.ed_name.text().strip()

    def marker_color(self) -> QColor:
        return self._color

# ───────────────────────────────────────────────────────────────
# Marker Manager popup

class MarkerManagerDialog(QDialog):
    """
    Popup to manage markers for the Summary tab.
    Shows marker name, colour, SN count and lets user:
      - Add marker
      - Assign SNs
      - Remove SNs
      - Delete marker
    """
    def __init__(self, summary_tab, parent=None):
        super().__init__(parent or summary_tab)
        self._tab = summary_tab
        self.setWindowTitle("Manage Markers")
        self.resize(520, 360)

        main = QVBoxLayout(self)
        main.addWidget(QLabel("Markers for Summary table:"))

        body = QHBoxLayout()
        self.lst_markers = QListWidget()
        self.lst_markers.setSelectionMode(QAbstractItemView.SingleSelection)
        body.addWidget(self.lst_markers, 1)

        btn_col = QVBoxLayout()
        self.btn_new = QPushButton("New Marker…")
        self.btn_assign = QPushButton("Assign SNs…")
        self.btn_remove = QPushButton("Remove SNs…")
        self.btn_delete = QPushButton("Delete Marker")
        self.btn_close = QPushButton("Close")

        for b in (self.btn_new, self.btn_assign, self.btn_remove, self.btn_delete, self.btn_close):
            btn_col.addWidget(b)
        btn_col.addStretch(1)
        body.addLayout(btn_col)

        main.addLayout(body)

        self.btn_new.clicked.connect(self._on_new)
        self.btn_assign.clicked.connect(self._on_assign)
        self.btn_remove.clicked.connect(self._on_remove)
        self.btn_delete.clicked.connect(self._on_delete)
        self.btn_close.clicked.connect(self.accept)

        self._refresh_list()

    def _refresh_list(self):
        self.lst_markers.clear()
        for name, info in self._tab._markers.items():
            sns = info.get("sns", set())
            color = info.get("color", QColor("#FFF3CD"))
            item = QListWidgetItem(f"{name} ({len(sns)} SNs)")
            item.setData(Qt.UserRole, name)
            item.setBackground(QBrush(color))
            self.lst_markers.addItem(item)

    def _selected_marker_name(self):
        it = self.lst_markers.currentItem()
        return it.data(Qt.UserRole) if it else None

    def _on_new(self):
        self._tab.add_marker(parent=self)
        self._refresh_list()

    def _on_assign(self):
        name = self._selected_marker_name()
        if not name:
            QMessageBox.information(self, "No marker selected", "Select a marker first.")
            return
        self._tab.assign_sns_to_marker(name, parent=self)
        self._refresh_list()

    def _on_remove(self):
        name = self._selected_marker_name()
        if not name:
            QMessageBox.information(self, "No marker selected", "Select a marker first.")
            return
        self._tab.remove_sns_from_marker(name, parent=self)
        self._refresh_list()

    def _on_delete(self):
        name = self._selected_marker_name()
        if not name:
            QMessageBox.information(self, "No marker selected", "Select a marker first.")
            return
        self._tab.delete_marker(name, parent=self)
        self._refresh_list()

# ───────────────────────────────────────────────────────────────
# Summary Tab

class SummaryTab(QWidget):
    def __init__(self):
        super().__init__()
        self._latest_type_to_frames = {}
        self._raw_type_to_frames = {}
        self._latest_components = []
        self._latest_wo_map = {}
        self._latest_summary_df = pd.DataFrame()
        self._excel_frames = {}
        self._excel_components_cache = []

        # marker state: {marker_name: {"color": QColor, "sns": set(str)}}
        self._markers: dict[str, dict] = {}

        # Cached inputs for the selected mode (so user enters IDs BEFORE Get Data)
        self._sql_mode_is_sn = None        # True => Search SN, False => Search WO / Fetch WO
        self._sql_sn_list = None           # list[str] or None
        self._sql_wo_lines = None          # list[str] or None
        self._excel_sn_selection = None    # None => not asked; [] => all; list[str] => specific SNs

        lay = QVBoxLayout(self)


        # Excel picker row
        line2 = QHBoxLayout()
        self.btn_pick_excel = QPushButton("Choose Excel…")
        self.btn_pick_excel.setVisible(False)
        self.lbl_excel_file = QLabel("")
        self.lbl_excel_file.setStyleSheet("color:#6B7280;")
        self.lbl_excel_file.setVisible(False)
        line2.addWidget(self.btn_pick_excel)
        line2.addWidget(self.lbl_excel_file, 1)
        lay.addLayout(line2)

        # ONE LINE: Filter + markers + main buttons
                # ROW 1: Source + Mode + Product + Filter + Markers
        row1 = QHBoxLayout()

        # Source dropdown (no label, as in your code)
        self.cbo_source = QComboBox()
        self.cbo_source.addItems(["SQL (default)", "Excel workbook"])
        row1.addWidget(self.cbo_source)

        row1.addSpacing(16)
        row1.addWidget(QLabel("Mode:"))
        self.cbo_mode = QComboBox()
        self.cbo_mode.addItems(["Search SN", "Search WO", "Fetch WO"])
        row1.addWidget(self.cbo_mode)

        row1.addSpacing(16)
        row1.addWidget(QLabel("Product:"))
        self.cb_dr8 = QCheckBox("DR8+")
        self.cb_fr4 = QCheckBox("FR4")
        row1.addWidget(self.cb_dr8)
        row1.addWidget(self.cb_fr4)

        row1.addSpacing(16)

        # Filter combo (now part of ROW 1)
        self.cbo_filter = QComboBox()
        self.cbo_filter.addItems([
            "All devices",
            "Only PASS devices",
            "Only FAIL devices",
            "All (exclude marked)",
            "Only PASS (exclude marked)",
            "Only FAIL (exclude marked)",
        ])
        row1.addWidget(self.cbo_filter)

        # Markers button (also ROW 1)
        self.btn_manage_markers = QPushButton("Markers…")
        row1.addWidget(self.btn_manage_markers)

        row1.addStretch(1)
        lay.addLayout(row1)

        # Excel picker row (unchanged)
        line2 = QHBoxLayout()
        self.btn_pick_excel = QPushButton("Choose Excel…")
        self.btn_pick_excel.setVisible(False)
        self.lbl_excel_file = QLabel("")
        self.lbl_excel_file.setStyleSheet("color:#6B7280;")
        self.lbl_excel_file.setVisible(False)
        line2.addWidget(self.btn_pick_excel)
        line2.addWidget(self.lbl_excel_file, 1)
        lay.addLayout(line2)

        # ROW 2: Get Data + Save All (left side)
        row2 = QHBoxLayout()
        self.btn_preview = QPushButton("Get Data")
        self.btn_save_all = QPushButton("Save All ")
        self.btn_save_all.setEnabled(False)
        row2.addWidget(self.btn_preview)
        row2.addWidget(self.btn_save_all)
        row2.addStretch(1)
        lay.addLayout(row2)

        # Progress
        self.prog = QProgressBar()
        self.prog.setRange(0, 100)
        self.prog.setVisible(False)
        lay.addWidget(self.prog)

        # Table
        self.tbl_model = DataFrameModel(pd.DataFrame())
        self.tbl_view  = QTableView()
        self.tbl_view.setModel(self.tbl_model)
        compact_table(self.tbl_view, row_h=26, min_col_w=110, first_col_w=140)
        lay.addWidget(self.tbl_view)
        # Double-click on a station cell → open detail popup
        self.tbl_view.doubleClicked.connect(self._on_summary_cell_double_clicked)



        # Signals
        self.cbo_source.currentIndexChanged.connect(self._on_source_changed)
        self.btn_pick_excel.clicked.connect(self._pick_excel)
        self.cb_dr8.toggled.connect(self._on_product_toggled)
        self.cb_fr4.toggled.connect(self._on_product_toggled)
        self.btn_preview.clicked.connect(self.refresh_now)
        self.btn_save_all.clicked.connect(self.save_all_excel)
        self.cbo_filter.currentIndexChanged.connect(self._reapply_filters)
        self.btn_manage_markers.clicked.connect(self._open_marker_manager)
        # NEW: popup as soon as mode is selected
        self.cbo_mode.activated.connect(self._on_mode_changed)

        self._on_source_changed(self.cbo_source.currentIndex())


    # UI helpers
    def _on_source_changed(self, idx: int):
        excel = (idx == 1)
        self.btn_pick_excel.setVisible(excel)
        self.lbl_excel_file.setVisible(excel)
        # reset cached inputs whenever source changes
        self._sql_mode_is_sn = None
        self._sql_sn_list = None
        self._sql_wo_lines = None
        self._excel_sn_selection = None

    def _on_product_toggled(self):
        sender = self.sender()
        if sender is self.cb_dr8 and self.cb_dr8.isChecked():
            self.cb_fr4.setChecked(False)
        elif sender is self.cb_fr4 and self.cb_fr4.isChecked():
            self.cb_dr8.setChecked(False)
        # if user is in Fetch WO mode and changes product, cached WOs are no longer valid
        if self.cbo_mode.currentText() == "Fetch WO":
            self._sql_wo_lines = None

    def _on_mode_changed(self, idx=None):
        """
        When user selects Search SN / Search WO / Fetch WO, show the popup
        immediately and cache the values. Get Data will just reuse these.
        """
        mode_text = self.cbo_mode.currentText()
        use_excel = (self.cbo_source.currentIndex() == 1)

        # Clear caches for this source
        if use_excel:
            self._excel_sn_selection = None
        else:
            self._sql_mode_is_sn = None
            self._sql_sn_list = None
            self._sql_wo_lines = None

        # Excel: only Search SN is allowed
        if use_excel and mode_text != "Search SN":
            QMessageBox.warning(
                self,
                "Excel mode",
                "In Excel source only 'Search SN' is supported.\n"
                "Reverting mode to 'Search SN'.",
            )
            self.cbo_mode.blockSignals(True)
            self.cbo_mode.setCurrentIndex(0)  # Search SN
            self.cbo_mode.blockSignals(False)
            return

        # Popups per mode, but no SQL yet
        if mode_text == "Search SN":
            if use_excel:
                dlg = InputPopup(
                    title="Search SN (Excel)",
                    label="Paste COMPONENTIDs (SN), one per line.\n"
                          "Leave empty to use all SNs in the workbook:",
                    parent=self,
                )
                if dlg.exec_() == QDialog.Accepted:
                    self._excel_sn_selection = dlg.lines()  # [] means "all"
            else:
                dlg = InputPopup(
                    title="Search SN (SQL)",
                    label="Paste COMPONENTIDs (SN), one per line:",
                    parent=self,
                )
                if dlg.exec_() == QDialog.Accepted:
                    lines = dlg.lines()
                    if lines:
                        self._sql_mode_is_sn = True
                        self._sql_sn_list = lines

        elif mode_text == "Search WO" and not use_excel:
            dlg = InputPopup(
                title="Search WO (SQL)",
                label="Paste Work Orders, one per line:",
                parent=self,
            )
            if dlg.exec_() == QDialog.Accepted:
                lines = dlg.lines()
                if lines:
                    self._sql_mode_is_sn = False
                    self._sql_wo_lines = lines

        elif mode_text == "Fetch WO" and not use_excel:
            chosen = self._fetch_wos_via_device()
            if chosen:
                self._sql_mode_is_sn = False
                self._sql_wo_lines = chosen

    # Busy
    def _start_busy(self):
        QApplication.setOverrideCursor(Qt.WaitCursor)
        self.prog.setVisible(True)
        self.prog.setValue(0)
        QApplication.processEvents()

    def _set_progress(self, pct: int):
        self.prog.setValue(max(0, min(100, pct)))
        QApplication.processEvents()

    def _end_busy(self):
        self.prog.setVisible(False)
        QApplication.restoreOverrideCursor()
        QApplication.processEvents()

    # Excel picker
    def _pick_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Excel", "", "Excel files (*.xlsx *.xlsm *.xls)")
        if not path:
            return
        try:
            self._start_busy()
            frames = pd.read_excel(path, sheet_name=None)
            frames = {k: pd.DataFrame(v) for k, v in frames.items()}
            self._excel_frames = frames
            self.lbl_excel_file.setText(os.path.basename(path))
            comps = set()
            for df in frames.values():
                if "COMPONENTID" in df.columns:
                    comps.update(str(x).strip() for x in df["COMPONENTID"].dropna().astype(str))
            self._excel_components_cache = sorted(c for c in comps if c)
        except Exception as e:
            QMessageBox.critical(self, "Excel Load Error", f"Failed to load workbook:\n{e}")
            self._excel_frames = {}
            self._excel_components_cache = []
            self.lbl_excel_file.setText("")
        finally:
            self._end_busy()

    # Fetch WOs via Device (DR8+/FR4) → returns list or None
    def _fetch_wos_via_device(self):
        if self.cbo_source.currentIndex() == 1:
            QMessageBox.information(self, "SQL Source Required",
                                    "Fetch WO via Device works only in SQL mode.")
            return None

        device_id = None
        if self.cb_dr8.isChecked():
            device_id = DEVICE_ID_DR8
        elif self.cb_fr4.isChecked():
            device_id = DEVICE_ID_FR4

        if not device_id:
            QMessageBox.warning(self, "Select Product",
                                "Check DR8+ or FR4 first before using Fetch WO.")
            return None

        try:
            self._start_busy()
            conn = pyodbc.connect(DB_CONN)
            wos = fetch_wos_for_device(conn, device_id)
            conn.close()
        except Exception as e:
            self._end_busy()
            QMessageBox.critical(self, "DB Error", f"Failed to fetch WOs:\n{e}")
            return None

        self._end_busy()

        if not wos:
            QMessageBox.information(self, "No WOs Found",
                                    f"No Work Orders found for Device {device_id}.")
            return None

        dlg = WoPickerDialog(wos, self)
        if dlg.exec_() != QDialog.Accepted:
            return None
        chosen = dlg.selected_wos()
        return chosen or None

    # Actions
    def refresh_now(self):
        use_excel = (self.cbo_source.currentIndex() == 1)
        mode_text = self.cbo_mode.currentText()

        self._start_busy()
        self.btn_preview.setEnabled(False)
        self.btn_save_all.setEnabled(False)

        # Excel mode: only Search SN is supported
        if use_excel:
            try:
                if not self._excel_frames:
                    raise RuntimeError("Choose an Excel workbook first.")

                if mode_text != "Search SN":
                    raise RuntimeError("In Excel mode only 'Search SN' is supported.")

                # Use cached SNs from mode selection, or ask now if not yet provided
                if self._excel_sn_selection is None:
                    dlg = InputPopup(
                        title="Search SN (Excel)",
                        label="Paste COMPONENTIDs (SN), one per line.\n"
                              "Leave empty to use all SNs in the workbook:",
                        parent=self,
                    )
                    if dlg.exec_() != QDialog.Accepted:
                        return
                    self._excel_sn_selection = dlg.lines()

                pasted = self._excel_sn_selection or []
                components = pasted if pasted else self._excel_components_cache
                if not components:
                    raise RuntimeError("No COMPONENTIDs found (paste SNs or provide a workbook with COMPONENTID columns).")

                self._set_progress(15)
                type_to_frames = parse_excel_to_type_frames(self._excel_frames)
                raw_type_to_frames = parse_excel_to_type_frames(self._excel_frames, reduce=False)
                self._set_progress(30)
                wo_map = extract_wo_map_from_excel(self._excel_frames, components)
                self._set_progress(55)
                summary_df = build_summary(type_to_frames, components, wo_map)

                self._latest_type_to_frames = type_to_frames
                self._raw_type_to_frames = raw_type_to_frames
                self._latest_components = components
                self._latest_wo_map = wo_map
                self._latest_summary_df = summary_df

                self._update_marker_colors_on_model()
                self._reapply_filters()

                self.btn_save_all.setEnabled(not summary_df.empty)
                self._set_progress(100)
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))
            finally:
                self._end_busy()
                self.btn_preview.setEnabled(True)
            return

        # SQL mode
        mode_is_sn = None
        sn_list = []
        wo_lines = []

        try:
            if mode_text == "Search SN":
                mode_is_sn = True
                if self._sql_sn_list is None:
                    dlg = InputPopup(
                        title="Search SN (SQL)",
                        label="Paste COMPONENTIDs (SN), one per line:",
                        parent=self,
                    )
                    if dlg.exec_() != QDialog.Accepted:
                        return
                    lines = dlg.lines()
                    if not lines:
                        raise RuntimeError("Please paste at least one SN.")
                    self._sql_mode_is_sn = True
                    self._sql_sn_list = lines
                sn_list = self._sql_sn_list

            elif mode_text == "Search WO":
                mode_is_sn = False
                if self._sql_wo_lines is None:
                    dlg = InputPopup(
                        title="Search WO (SQL)",
                        label="Paste Work Orders, one per line:",
                        parent=self,
                    )
                    if dlg.exec_() != QDialog.Accepted:
                        return
                    lines = dlg.lines()
                    if not lines:
                        raise RuntimeError("Please paste at least one Work Order.")
                    self._sql_mode_is_sn = False
                    self._sql_wo_lines = lines
                wo_lines = self._sql_wo_lines

            elif mode_text == "Fetch WO":
                mode_is_sn = False
                if self._sql_wo_lines is None:
                    chosen = self._fetch_wos_via_device()
                    if not chosen:
                        # user cancelled or nothing selected
                        self._end_busy()
                        self.btn_preview.setEnabled(True)
                        return
                    self._sql_mode_is_sn = False
                    self._sql_wo_lines = chosen
                wo_lines = self._sql_wo_lines

            else:
                raise RuntimeError(f"Unsupported mode: {mode_text}")

            restrict_device_id = None


            if mode_is_sn:


                if self.cb_dr8.isChecked():


                    restrict_device_id = DEVICE_ID_DR8


                elif self.cb_fr4.isChecked():


                    restrict_device_id = DEVICE_ID_FR4


            self.worker = QueryWorker(mode_is_sn, sn_list, wo_lines, build_compare=False, restrict_device_id=restrict_device_id)
            self.worker.progressPct.connect(self._set_progress)
            self.worker.done.connect(self._worker_done_from_sql)
            self.worker.error.connect(self._worker_error)
            self.worker.start()

        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))
            self._end_busy()
            self.btn_preview.setEnabled(True)

    def _worker_done_from_sql(self, type_to_frames, comps, wo_map, summary_df, _ddmi_live_df, _ddmi_blocks):
        self._latest_type_to_frames = type_to_frames
        self._raw_type_to_frames = getattr(self.worker, '_raw_type_to_frames', {})
        self._latest_components = comps
        self._latest_wo_map = wo_map
        self._latest_summary_df = summary_df

        self._update_marker_colors_on_model()
        self._reapply_filters()

        self.btn_preview.setEnabled(True)
        self.btn_save_all.setEnabled(not summary_df.empty)
        self._end_busy()

    def _worker_error(self, msg: str):
        QMessageBox.critical(self, "Error", msg)
        self.btn_preview.setEnabled(True)
        self._end_busy()

    # Live WO merging (same WO in one merged cell)
    def _apply_wo_spans(self):
        self.tbl_view.clearSpans()
        df = self.tbl_model._df if hasattr(self.tbl_model, "_df") else self._latest_summary_df
        if df is None or df.empty:
            return
        cols = list(df.columns)
        if "WO" not in cols:
            return
        wo_col = cols.index("WO")
        row_count = len(df)
        r = 0
        while r < row_count:
            val = df.iloc[r, wo_col]
            if val in ("", None):
                r += 1
                continue
            start = r
            r += 1
            while r < row_count and df.iloc[r, wo_col] == val:
                r += 1
            span_len = r - start
            if span_len > 1:
                self.tbl_view.setSpan(start, wo_col, span_len, 1)
    def _on_summary_cell_double_clicked(self, index):
        """
        When user double-clicks a cell in the SUMMARY table:
        - Identify SN (row) and station (column)
        - For normal stations: show raw rows in a simple table.
        - For 3TBER / TCBER / Final Test: show a multi-tab popup with RT/LT/HT.
        """
        if not index.isValid():
            return

        df_view = self.tbl_model._df
        if df_view is None or df_view.empty:
            return

        row = index.row()
        col = index.column()

        # Station name (column header)
        try:
            col_name = str(df_view.columns[col])
        except Exception:
            return

        # Ignore WO / SN header columns
        if col_name in ("WO", "SN"):
            return

        # Empty cell → nothing to show
        cell_val = df_view.iat[row, col]
        if pd.isna(cell_val) or str(cell_val).strip() == "":
            return

        # SN for this row
        try:
            sn = str(df_view.loc[row, "SN"]).strip()
        except Exception:
            return
        if not sn:
            return

        # Map summary column → (family, subtype)
        col_map = dict(SUMMARY_COLUMNS)
        if col_name not in col_map:
            return
        family, subtype = col_map[col_name]

        # We need the raw frame corresponding to this station
        if not self._latest_type_to_frames:
            QMessageBox.information(
                self, "No raw data",
                "Raw test data is not available for this query.\n"
                "Click 'Get Data' first."
            )
            return

        try:
            if family == "TRX":
                trx_map = self._latest_type_to_frames.get("TRX", {})
                df_all = trx_map.get(subtype, pd.DataFrame())
            else:
                df_all = self._latest_type_to_frames.get(family, pd.DataFrame())
        except Exception:
            df_all = pd.DataFrame()

        if df_all is None or df_all.empty:
            QMessageBox.information(
                self, "No data",
                f"No raw '{col_name}' rows found for SN {sn}."
            )
            return

        # Filter to this single SN
        try:
            sub = df_all[df_all["COMPONENTID"].astype(str).str.strip() == sn].copy()
        except Exception:
            QMessageBox.information(
                self, "No data",
                f"No raw '{col_name}' rows found for SN {sn}."
            )
            return

        if sub.empty:
            QMessageBox.information(
                self, "No data",
                f"No raw '{col_name}' rows found for SN {sn}."
            )
            return

        # Look up unreduced (all-SID) data for the SID selector dropdown
        try:
            if family == "TRX":
                raw_trx_map = self._raw_type_to_frames.get("TRX", {})
                raw_df_all = raw_trx_map.get(subtype, pd.DataFrame())
            else:
                raw_df_all = self._raw_type_to_frames.get(family, pd.DataFrame())
        except Exception:
            raw_df_all = pd.DataFrame()
        # Filter raw data to this SN
        if raw_df_all is not None and not raw_df_all.empty:
            try:
                sub_raw = raw_df_all[raw_df_all["COMPONENTID"].astype(str).str.strip() == sn].copy()
            except Exception:
                sub_raw = pd.DataFrame()
        else:
            sub_raw = pd.DataFrame()

        # Apply DDMI gate so detail matches SUMMARY behaviour
        gate_sid = None
        try:
            trx_map = self._latest_type_to_frames.get("TRX", {})
            ddmi_gate_map = compute_ddmi_gate_sid_for_components(trx_map, [sn])
            gate_sid = ddmi_gate_map.get(sn)
        except Exception:
            gate_sid = None

        if gate_sid is not None and "_SID_NUM_" in sub.columns and col_name not in ("DDMI Cal", "FW Writing", "Mode Hopping"):
            sid_num = pd.to_numeric(sub["_SID_NUM_"], errors="coerce")
            sub = sub[sid_num >= gate_sid].copy()
            # Also gate the raw data
            if not sub_raw.empty and "_SID_NUM_" in sub_raw.columns:
                raw_sid_num = pd.to_numeric(sub_raw["_SID_NUM_"], errors="coerce")
                sub_raw = sub_raw[raw_sid_num >= gate_sid].copy()
            if sub.empty:
                QMessageBox.information(
                    self, "No data after gating",
                    f"All rows for SN {sn} at '{col_name}' were filtered out by DDMI gate."
                )
                return

        # ── SPECIAL CASE: 3TBER / TCBER / Final Test → 3-temp pivot + raw ────
        if col_name in ("3TBER", "TCBER", "Final Test"):
            dlg = ThreeTempPivotDialog(sn, col_name, sub, raw_all_df=sub_raw, parent=self)
            dlg.exec_()
            return


        # ── DEFAULT: simple one-table popup for other stations ───────────────
        dlg = StationDetailDialog(sn, col_name, sub, raw_all_df=sub_raw, parent=self)
        dlg.exec_()

    # Marker helpers
    def _compute_marker_color_map(self) -> dict:
        mapping = {}
        for name, info in self._markers.items():
            color = info.get("color", None)
            sns = info.get("sns", set())
            if not isinstance(color, QColor):
                continue
            for sn in sns:
                key = str(sn).strip()
                if key:
                    mapping[key] = color  # last marker wins if overlaps
        return mapping

    def _update_marker_colors_on_model(self):
        mapping = self._compute_marker_color_map()
        self.tbl_model.setMarkers(mapping)

    def _apply_marker_filter(self, df: pd.DataFrame, exclude_marked: bool) -> pd.DataFrame:
        if df is None or df.empty or not exclude_marked:
            return df
        if not self._markers or "SN" not in df.columns:
            return df
        marked = set()
        for info in self._markers.values():
            marked |= {str(s).strip() for s in info.get("sns", set()) if str(s).strip()}
        if not marked:
            return df
        mask = ~df["SN"].astype(str).str.strip().isin(marked)
        return df[mask].reset_index(drop=True)

    def _reapply_filters(self):
        """Reapply PASS/FAIL + marker exclude filter on latest_summary_df."""
        if self._latest_summary_df is None or self._latest_summary_df.empty:
            self.tbl_model.setDataFrame(pd.DataFrame())
            return

        df = self._latest_summary_df.copy()

        text = self.cbo_filter.currentText() if hasattr(self, "cbo_filter") else "All devices"
        t = text.lower()

        pass_only = "only pass" in t
        fail_only = "only fail" in t
        exclude_marked = "exclude marked" in t

        # PASS/FAIL logic (same as older code but driven by dropdown)
        if pass_only or fail_only:
            test_cols = [c for c in df.columns if c not in ("WO", "SN")]
            if test_cols:
                vals = df[test_cols].apply(
                    lambda col: col.fillna("").astype(str).str.strip().str.upper()
                )
                if pass_only:
                    mask = vals.apply(lambda row: all(v in ("", "PASS") for v in row), axis=1)
                else:
                    mask = vals.apply(lambda row: any(v not in ("", "PASS") for v in row), axis=1)
                df = df[mask].reset_index(drop=True)

        df = self._apply_marker_filter(df, exclude_marked)
        self.tbl_model.setDataFrame(df)
        self._update_marker_colors_on_model()
        compact_table(self.tbl_view)
        self._apply_wo_spans()

    # Marker API (used by MarkerManagerDialog)
    def _open_marker_manager(self):
        dlg = MarkerManagerDialog(self, self)
        dlg.exec_()
        self._update_marker_colors_on_model()
        self._reapply_filters()

    def add_marker(self, parent=None):
        dlg = MarkerDialog(existing_names=list(self._markers.keys()), parent=parent or self)
        if dlg.exec_() != QDialog.Accepted:
            return None
        name = dlg.marker_name()
        color = dlg.marker_color()
        if not name:
            return None
        self._markers[name] = {"color": color, "sns": set()}
        self._update_marker_colors_on_model()
        self._reapply_filters()
        return name

    def assign_sns_to_marker(self, marker_name, parent=None):
        if marker_name not in self._markers:
            QMessageBox.information(self, "Marker not found", "Please select a valid marker.")
            return
        dlg = InputPopup(
            title=f"Assign SNs to '{marker_name}'",
            label="Paste SNs (one per line) to assign:",
            parent=parent or self,
        )
        if dlg.exec_() != QDialog.Accepted:
            return
        sn_list = dlg.lines()
        if not sn_list:
            return
        sn_set = {str(s).strip() for s in sn_list if str(s).strip()}
        if self._latest_summary_df is not None and not self._latest_summary_df.empty and "SN" in self._latest_summary_df.columns:
            valid_sns = set(self._latest_summary_df["SN"].astype(str).str.strip())
            sn_set = {s for s in sn_set if s in valid_sns}
        if not sn_set:
            QMessageBox.information(self, "No matches", "None of the pasted SNs were found in the current SUMMARY table.")
            return
        self._markers[marker_name]["sns"].update(sn_set)
        self._update_marker_colors_on_model()
        self._reapply_filters()

    def remove_sns_from_marker(self, marker_name, parent=None):
        if marker_name not in self._markers:
            QMessageBox.information(self, "Marker not found", "Please select a valid marker.")
            return
        dlg = InputPopup(
            title=f"Remove SNs from '{marker_name}'",
            label="Paste SNs (one per line) to remove:",
            parent=parent or self,
        )
        if dlg.exec_() != QDialog.Accepted:
            return
        sn_list = dlg.lines()
        if not sn_list:
            return
        sn_set = {str(s).strip() for s in sn_list if str(s).strip()}
        self._markers[marker_name]["sns"] -= sn_set
        self._update_marker_colors_on_model()
        self._reapply_filters()

    def delete_marker(self, marker_name, parent=None):
        if marker_name not in self._markers:
            return
        reply = QMessageBox.question(
            parent or self, "Delete marker",
            f"Delete marker '{marker_name}' and remove its highlights?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return
        self._markers.pop(marker_name, None)
        self._update_marker_colors_on_model()
        self._reapply_filters()

    def save_all_excel(self):
        if self._latest_summary_df is None or self._latest_summary_df.empty:
            QMessageBox.information(self, "Nothing to Save", "No SUMMARY to save. Click Get Data first.")
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self, "Save Excel File", "latest_test_results.xlsx", "Excel Files (*.xlsx)"
        )
        if not save_path:
            return

        try:
            from openpyxl.styles import PatternFill, Alignment

            self._start_busy()
            type_to_frames = self._latest_type_to_frames
            trx_map = type_to_frames.get("TRX", {}) if isinstance(type_to_frames.get("TRX", {}), dict) else {}
            other_keys = [k for k in type_to_frames.keys() if k != "TRX"]

            with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                # raw TRX sub-sheets
                for k, df in trx_map.items():
                    if isinstance(df, pd.DataFrame) and not df.empty and \
                       k in {"DDMI", "RT", "LT", "HT", "FINAL", "FINAL_ATS", "FINAL_LT", "FINAL_HT"}:
                        df.to_excel(writer, index=False, sheet_name=k)

                # other sheets
                for k in other_keys:
                    df = type_to_frames[k]
                    if isinstance(df, pd.DataFrame) and not df.empty:
                        df.to_excel(writer, index=False, sheet_name=k)

                # SUMMARY
                summary_df = self._latest_summary_df.copy()
                summary_df.to_excel(writer, index=False, sheet_name="SUMMARY")

                wb = writer.book
                ws = wb["SUMMARY"]

                # Merge contiguous WO cells (column A)
                col_idx = 1
                r = 2
                max_row = ws.max_row
                while r <= max_row:
                    v = ws.cell(r, col_idx).value
                    if v in (None, "", " "):
                        r += 1
                        continue
                    start = r
                    while r + 1 <= max_row and ws.cell(r + 1, col_idx).value == v:
                        r += 1
                    if r > start:
                        ws.merge_cells(start_row=start, start_column=col_idx,
                                       end_row=r, end_column=col_idx)
                        ws.cell(start, col_idx).alignment = Alignment(
                            vertical="center", horizontal="left"
                        )
                    r += 1

                # Conditional colouring: PASS green, others red (C..end)
                fill_pass = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                fill_fail = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

                max_row, max_col = ws.max_row, ws.max_column
                for rr in range(2, max_row + 1):
                    for cc in range(3, max_col + 1):
                        cell = ws.cell(row=rr, column=cc)
                        v = cell.value
                        v = v.strip() if isinstance(v, str) else v
                        if v == "PASS":
                            cell.fill = fill_pass
                        elif v not in (None, "", " "):
                            cell.fill = fill_fail

            QMessageBox.information(self, "Complete", f"Excel saved:\n{save_path}")

        except Exception as e:
            QMessageBox.critical(self, "Save Error", f"Failed to save:\n{e}")
        finally:
            self._end_busy()

# ───────────────────────────────────────────────────────────────
# Compare Tab

class CompareTab(QWidget):
    def __init__(self):
        super().__init__()
        self._latest_type_to_frames = {}
        self._raw_type_to_frames = {}
        self._latest_components = []
        self._latest_wo_map = {}
        self._latest_ddmi_live_df = pd.DataFrame()
        self._latest_ddmi_blocks  = []
        self._sel_metrics = DDMI_METRICS[:] + EXTRA_METRICS[:]
        self._sel_sheets  = SHEET_ORDER[:]
        self._excel_frames = {}
        self._excel_components_cache = []

        # Cached inputs for mode
        self._sql_mode_is_sn = None
        self._sql_sn_list = None
        self._sql_wo_lines = None
        self._excel_sn_selection = None

        lay = QVBoxLayout(self)

        # Source & Mode + Product
        line1 = QHBoxLayout()
        line1.addWidget(QLabel("Source:"))
        self.cbo_source = QComboBox()
        self.cbo_source.addItems(["SQL (default)", "Excel workbook"])
        line1.addWidget(self.cbo_source)
        line1.addSpacing(16)
        line1.addWidget(QLabel("Mode:"))
        self.cbo_mode = QComboBox()
        self.cbo_mode.addItems(["Search SN", "Search WO", "Fetch WO"])
        line1.addWidget(self.cbo_mode)
        line1.addSpacing(16)
        line1.addWidget(QLabel("Product:"))
        self.cb_dr8 = QCheckBox("DR8+")
        self.cb_fr4 = QCheckBox("FR4")
        line1.addWidget(self.cb_dr8)
        line1.addWidget(self.cb_fr4)
        line1.addStretch(1)
        lay.addLayout(line1)

        # Inputs (Excel)
        line2 = QHBoxLayout()
        self.btn_pick_excel = QPushButton("Choose Excel…")
        self.btn_pick_excel.setVisible(False)
        self.lbl_excel_file = QLabel("")
        self.lbl_excel_file.setStyleSheet("color:#6B7280;")
        self.lbl_excel_file.setVisible(False)
        line2.addWidget(self.btn_pick_excel)
        line2.addWidget(self.lbl_excel_file, 1)
        lay.addLayout(line2)

        # Global selection
        opt_box = QGroupBox("Global selection for Compare")
        opt_lay = QHBoxLayout(opt_box)
        self.lst_metrics = QListWidget()
        self.lst_metrics.setSelectionMode(QListWidget.MultiSelection)
        for m in (DDMI_METRICS + EXTRA_METRICS):
            it = QListWidgetItem(m)
            it.setSelected(True)
            self.lst_metrics.addItem(it)
        self.lst_sheets = QListWidget()
        self.lst_sheets.setSelectionMode(QListWidget.MultiSelection)
        for s in SHEET_ORDER:
            it = QListWidgetItem(s)
            it.setSelected(True)
            self.lst_sheets.addItem(it)
        self.cbo_preset = QComboBox()
        self.cbo_preset.addItems(["All", "TX/RX only", "Vcc only", "Optical Quality (OMA/ER/TDECQ/RLM)"])

        def _apply_preset(idx):
            choose_all = (idx == 0)
            txrx_only  = (idx == 1)
            vcc_only   = (idx == 2)
            oq         = (idx == 3)
            for i in range(self.lst_metrics.count()):
                txt = self.lst_metrics.item(i).text()
                sel = True
                if txrx_only:
                    sel = txt in ["Power(dBm)", "DDMI_TxP", "DD_RxP1", "dTxP", "dRxP1",
                                  "Txp_Offset", "Rxp_Offset", "Txp_Slope", "Rxp_Slope"]
                elif vcc_only:
                    sel = txt in ["I2C_Vcc", "DDMI_Vcc", "dVcc(%)"]
                elif oq:
                    sel = txt in ["Outer_OMA(dB)", "Outer_ER(dB)", "TDECQ(dB)", "OMA_TDECQ(dB)",
                                  "RLM", "TDECQ_Ceq(dB)", "Bias_Setpoint", "Tx_Disable(dBm)",
                                  "TDECQ_4.8E-6"]
                self.lst_metrics.item(i).setSelected(choose_all or sel)
            for i in range(self.lst_sheets.count()):
                self.lst_sheets.item(i).setSelected(True)

        self.cbo_preset.currentIndexChanged.connect(_apply_preset)
        self.btn_apply_global = QPushButton("Apply selection")
        self.btn_apply_global.clicked.connect(self._apply_current_selection)

        opt_lay.addWidget(QLabel("Presets:"))
        opt_lay.addWidget(self.cbo_preset)
        opt_lay.addSpacing(12)
        opt_lay.addWidget(QLabel("Metrics:"))
        opt_lay.addWidget(self.lst_metrics, 1)
        opt_lay.addWidget(QLabel("Sheets:"))
        opt_lay.addWidget(self.lst_sheets, 1)
        opt_lay.addStretch(1)
        opt_lay.addWidget(self.btn_apply_global)
        lay.addWidget(opt_box)

        # Controls
        ctrl = QHBoxLayout()
        self.btn_preview = QPushButton("Get Data")
        self.btn_save_cmp = QPushButton("Save Compare (4 sheets)…")
        self.btn_save_cmp.setEnabled(False)
        ctrl.addWidget(self.btn_preview)
        ctrl.addStretch(1)
        ctrl.addWidget(self.btn_save_cmp)
        lay.addLayout(ctrl)

        # Progress
        self.prog = QProgressBar()
        self.prog.setRange(0, 100)
        self.prog.setVisible(False)
        lay.addWidget(self.prog)

        # Tab-local filters
        filt_box = QGroupBox("Filters (local to this page)")
        f_lay = QHBoxLayout(filt_box)
        self.cbo_metric = QComboBox()
        self.cbo_metric.addItems(["All"] + (DDMI_METRICS + EXTRA_METRICS))
        self.cbo_sheet  = QComboBox()
        self.cbo_sheet.addItems(["All"] + SHEET_ORDER)
        self.btn_apply_tab = QPushButton("Apply")
        f_lay.addWidget(QLabel("Metric:"))
        f_lay.addWidget(self.cbo_metric)
        f_lay.addWidget(QLabel("Sheet:"))
        f_lay.addWidget(self.cbo_sheet)
        f_lay.addStretch(1)
        f_lay.addWidget(self.btn_apply_tab)
        lay.addWidget(filt_box)

        # Compare table
        self.ddmi_model = DataFrameModel(pd.DataFrame())
        self.ddmi_view  = QTableView()
        self.ddmi_view.setModel(self.ddmi_model)
        compact_table(self.ddmi_view, row_h=26, min_col_w=110, first_col_w=140)
        lay.addWidget(self.ddmi_view)

        # Signals
        self.cbo_source.currentIndexChanged.connect(self._on_source_changed)
        self.btn_pick_excel.clicked.connect(self._pick_excel)
        self.cb_dr8.toggled.connect(self._on_product_toggled)
        self.cb_fr4.toggled.connect(self._on_product_toggled)
        self.btn_preview.clicked.connect(self.refresh_now)
        self.btn_apply_tab.clicked.connect(self._apply_filters_tab)
        self.btn_save_cmp.clicked.connect(self.save_compare_excel)
        # NEW: popup on mode selection
        self.cbo_mode.activated.connect(self._on_mode_changed)

        self._on_source_changed(self.cbo_source.currentIndex())

    # UI helpers
    def _on_source_changed(self, idx: int):
        excel = (idx == 1)
        self.btn_pick_excel.setVisible(excel)
        self.lbl_excel_file.setVisible(excel)
        # reset cached inputs
        self._sql_mode_is_sn = None
        self._sql_sn_list = None
        self._sql_wo_lines = None
        self._excel_sn_selection = None

    def _on_product_toggled(self):
        sender = self.sender()
        if sender is self.cb_dr8 and self.cb_dr8.isChecked():
            self.cb_fr4.setChecked(False)
        elif sender is self.cb_fr4 and self.cb_fr4.isChecked():
            self.cb_dr8.setChecked(False)
        if self.cbo_mode.currentText() == "Fetch WO":
            self._sql_wo_lines = None

    def _on_mode_changed(self, idx=None):
        """
        Same idea as Summary tab: as soon as mode is selected, show popup
        and cache IDs, but only run query when Get Data is pressed.
        """
        mode_text = self.cbo_mode.currentText()
        use_excel = (self.cbo_source.currentIndex() == 1)

        if use_excel:
            self._excel_sn_selection = None
        else:
            self._sql_mode_is_sn = None
            self._sql_sn_list = None
            self._sql_wo_lines = None

        # Excel only supports Search SN
        if use_excel and mode_text != "Search SN":
            QMessageBox.warning(
                self,
                "Excel mode",
                "In Excel source only 'Search SN' is supported.\n"
                "Reverting mode to 'Search SN'.",
            )
            self.cbo_mode.blockSignals(True)
            self.cbo_mode.setCurrentIndex(0)
            self.cbo_mode.blockSignals(False)
            return

        if mode_text == "Search SN":
            if use_excel:
                dlg = InputPopup(
                    title="Search SN (Excel)",
                    label="Paste COMPONENTIDs (SN), one per line.\n"
                          "Leave empty to use all SNs in the workbook:",
                    parent=self,
                )
                if dlg.exec_() == QDialog.Accepted:
                    self._excel_sn_selection = dlg.lines()
            else:
                dlg = InputPopup(
                    title="Search SN (SQL)",
                    label="Paste COMPONENTIDs (SN), one per line:",
                    parent=self,
                )
                if dlg.exec_() == QDialog.Accepted:
                    lines = dlg.lines()
                    if lines:
                        self._sql_mode_is_sn = True
                        self._sql_sn_list = lines

        elif mode_text == "Search WO" and not use_excel:
            dlg = InputPopup(
                title="Search WO (SQL)",
                label="Paste Work Orders, one per line:",
                parent=self,
            )
            if dlg.exec_() == QDialog.Accepted:
                lines = dlg.lines()
                if lines:
                    self._sql_mode_is_sn = False
                    self._sql_wo_lines = lines

        elif mode_text == "Fetch WO" and not use_excel:
            chosen = self._fetch_wos_via_device()
            if chosen:
                self._sql_mode_is_sn = False
                self._sql_wo_lines = chosen

    # Busy
    def _start_busy(self):
        QApplication.setOverrideCursor(Qt.WaitCursor)
        self.prog.setVisible(True)
        self.prog.setValue(0)
        QApplication.processEvents()

    def _set_progress(self, pct: int):
        self.prog.setValue(max(0, min(100, pct)))
        QApplication.processEvents()

    def _end_busy(self):
        self.prog.setVisible(False)
        QApplication.restoreOverrideCursor()
        QApplication.processEvents()

    # Excel picker
    def _pick_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Excel", "", "Excel files (*.xlsx *.xlsm *.xls)")
        if not path:
            return
        try:
            self._start_busy()
            frames = pd.read_excel(path, sheet_name=None)
            frames = {k: pd.DataFrame(v) for k, v in frames.items()}
            self._excel_frames = frames
            self.lbl_excel_file.setText(os.path.basename(path))
            comps = set()
            for df in frames.values():
                if "COMPONENTID" in df.columns:
                    comps.update(str(x).strip() for x in df["COMPONENTID"].dropna().astype(str))
            self._excel_components_cache = sorted(c for c in comps if c)
        except Exception as e:
            QMessageBox.critical(self, "Excel Load Error", f"Failed to load workbook:\n{e}")
            self._excel_frames = {}
            self._excel_components_cache = []
            self.lbl_excel_file.setText("")
        finally:
            self._end_busy()

    # Fetch WOs via Device
    def _fetch_wos_via_device(self):
        if self.cbo_source.currentIndex() == 1:
            QMessageBox.information(self, "SQL Source Required",
                                    "Fetch WO via Device works only in SQL mode.")
            return None

        device_id = None
        if self.cb_dr8.isChecked():
            device_id = DEVICE_ID_DR8
        elif self.cb_fr4.isChecked():
            device_id = DEVICE_ID_FR4

        if not device_id:
            QMessageBox.warning(self, "Select Product",
                                "Check DR8+ or FR4 first before using Fetch WO.")
            return None

        try:
            self._start_busy()
            conn = pyodbc.connect(DB_CONN)
            wos = fetch_wos_for_device(conn, device_id)
            conn.close()
        except Exception as e:
            self._end_busy()
            QMessageBox.critical(self, "DB Error", f"Failed to fetch WOs:\n{e}")
            return None

        self._end_busy()

        if not wos:
            QMessageBox.information(self, "No WOs Found",
                                    f"No Work Orders found for Device {device_id}.")
            return None

        dlg = WoPickerDialog(wos, self)
        if dlg.exec_() != QDialog.Accepted:
            return None
        chosen = dlg.selected_wos()
        return chosen or None

    def _apply_current_selection(self):
        metrics = [
            self.lst_metrics.item(i).text()
            for i in range(self.lst_metrics.count())
            if self.lst_metrics.item(i).isSelected()
        ]
        sheets  = [
            self.lst_sheets.item(i).text()
            for i in range(self.lst_sheets.count())
            if self.lst_sheets.item(i).isSelected()
        ]
        self._sel_metrics = metrics if metrics else (DDMI_METRICS + EXTRA_METRICS)
        self._sel_sheets  = sheets if sheets else SHEET_ORDER
        self._rebuild_compare_only()

    def refresh_now(self):
        use_excel = (self.cbo_source.currentIndex() == 1)
        mode_text = self.cbo_mode.currentText()

        self._start_busy()
        self.btn_preview.setEnabled(False)
        self.btn_save_cmp.setEnabled(False)

        # Excel mode: only Search SN
        if use_excel:
            try:
                if not self._excel_frames:
                    raise RuntimeError("Choose an Excel workbook first.")
                if mode_text != "Search SN":
                    raise RuntimeError("In Excel mode only 'Search SN' is supported.")

                if self._excel_sn_selection is None:
                    dlg = InputPopup(
                        title="Search SN (Excel)",
                        label="Paste COMPONENTIDs (SN), one per line (leave empty to use all SNs in workbook):",
                        parent=self,
                    )
                    if dlg.exec_() != QDialog.Accepted:
                        return
                    self._excel_sn_selection = dlg.lines()

                pasted = self._excel_sn_selection or []
                components = pasted if pasted else self._excel_components_cache
                if not components:
                    raise RuntimeError("No COMPONENTIDs found (paste SNs or provide a workbook with COMPONENTID columns).")

                self._set_progress(10)
                type_to_frames = parse_excel_to_type_frames(self._excel_frames)
                raw_type_to_frames = parse_excel_to_type_frames(self._excel_frames, reduce=False)
                self._set_progress(25)
                wo_map = extract_wo_map_from_excel(self._excel_frames, components)
                self._set_progress(45)
                live_df, blocks = build_ddmi_compare_live_and_blocks(
                    type_to_frames, components, wo_map,
                    metrics=self._sel_metrics, sheets=self._sel_sheets
                )
                self._latest_type_to_frames = type_to_frames
                self._raw_type_to_frames = raw_type_to_frames
                self._latest_components = components
                self._latest_wo_map = wo_map
                self._latest_ddmi_live_df = live_df
                self._latest_ddmi_blocks = blocks
                self.ddmi_model.setDataFrame(live_df)
                compact_table(self.ddmi_view)
                self.btn_save_cmp.setEnabled(not live_df.empty)
                self._set_progress(100)
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))
            finally:
                self._end_busy()
                self.btn_preview.setEnabled(True)
            return

        # SQL path
        mode_is_sn = None
        sn_list = []
        wo_lines = []

        try:
            if mode_text == "Search SN":
                mode_is_sn = True
                if self._sql_sn_list is None:
                    dlg = InputPopup(
                        title="Search SN (SQL)",
                        label="Paste COMPONENTIDs (SN), one per line:",
                        parent=self,
                    )
                    if dlg.exec_() != QDialog.Accepted:
                        return
                    lines = dlg.lines()
                    if not lines:
                        raise RuntimeError("Please paste at least one SN.")
                    self._sql_mode_is_sn = True
                    self._sql_sn_list = lines
                sn_list = self._sql_sn_list

            elif mode_text == "Search WO":
                mode_is_sn = False
                if self._sql_wo_lines is None:
                    dlg = InputPopup(
                        title="Search WO (SQL)",
                        label="Paste Work Orders, one per line:",
                        parent=self,
                    )
                    if dlg.exec_() != QDialog.Accepted:
                        return
                    lines = dlg.lines()
                    if not lines:
                        raise RuntimeError("Please paste at least one Work Order.")
                    self._sql_mode_is_sn = False
                    self._sql_wo_lines = lines
                wo_lines = self._sql_wo_lines

            elif mode_text == "Fetch WO":
                mode_is_sn = False
                if self._sql_wo_lines is None:
                    chosen = self._fetch_wos_via_device()
                    if not chosen:
                        self._end_busy()
                        self.btn_preview.setEnabled(True)
                        return
                    self._sql_mode_is_sn = False
                    self._sql_wo_lines = chosen
                wo_lines = self._sql_wo_lines

            else:
                raise RuntimeError(f"Unsupported mode: {mode_text}")

            restrict_device_id = None


            if mode_is_sn:


                if self.cb_dr8.isChecked():


                    restrict_device_id = DEVICE_ID_DR8


                elif self.cb_fr4.isChecked():


                    restrict_device_id = DEVICE_ID_FR4


            self.worker = QueryWorker(mode_is_sn, sn_list, wo_lines, build_compare=True, restrict_device_id=restrict_device_id)
            self.worker.progressPct.connect(self._set_progress)
            self.worker.done.connect(self._worker_done_from_sql)
            self.worker.error.connect(self._worker_error)
            self.worker.start()

        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))
            self._end_busy()
            self.btn_preview.setEnabled(True)

    def _worker_done_from_sql(self, type_to_frames, comps, wo_map, _summary_df, ddmi_live_df, ddmi_blocks):
        self._latest_type_to_frames = type_to_frames
        self._raw_type_to_frames = getattr(self.worker, '_raw_type_to_frames', {})
        self._latest_components = comps
        self._latest_wo_map = wo_map
        self._latest_ddmi_live_df = ddmi_live_df
        self._latest_ddmi_blocks = ddmi_blocks
        self.ddmi_model.setDataFrame(ddmi_live_df)
        compact_table(self.ddmi_view)
        self.btn_preview.setEnabled(True)
        self.btn_save_cmp.setEnabled(not ddmi_live_df.empty)
        self._end_busy()

    def _worker_error(self, msg: str):
        QMessageBox.critical(self, "Error", msg)
        self.btn_preview.setEnabled(True)
        self._end_busy()

    def _rebuild_compare_only(self):
        if not self._latest_components or not self._latest_type_to_frames:
            return
        live_df, blocks = build_ddmi_compare_live_and_blocks(
            self._latest_type_to_frames, self._latest_components, self._latest_wo_map,
            metrics=self._sel_metrics, sheets=self._sel_sheets
        )
        self._latest_ddmi_live_df = live_df
        self._latest_ddmi_blocks = blocks
        self.ddmi_model.setDataFrame(live_df)
        compact_table(self.ddmi_view)

    def _apply_filters_tab(self):
        if self._latest_ddmi_live_df.empty:
            return
        df = self._latest_ddmi_live_df.copy()
        m = self.cbo_metric.currentText()
        s = self.cbo_sheet.currentText()
        if m != "All":
            df = df[df["Metric"] == m]
        if s != "All":
            base_cols = ["WO", "SN", "Metric", "CH"]
            keep = [c for c in (base_cols + [s, "StdAcross"]) if c in df.columns]
            df = df[keep]
        self.ddmi_model.setDataFrame(df)
        compact_table(self.ddmi_view)

    def save_compare_excel(self):
        if self._latest_ddmi_live_df is None or self._latest_ddmi_live_df.empty:
            QMessageBox.information(self, "Nothing to Save", "No Compare table to save. Click Get Data first.")
            return
        save_path, _ = QFileDialog.getSaveFileName(
            self, "Save Compare Excel", "ddmi_compare.xlsx", "Excel Files (*.xlsx)"
        )
        if not save_path:
            return
        try:
            self._start_busy()
            with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                wb = writer.book
                render_ddmi_compare_sheet(
                    wb, "DDMI_Compare",
                    self._latest_ddmi_blocks,
                    self._latest_components,
                    self._sel_metrics,
                    self._sel_sheets
                )
                for sheet_name, prof in [("Compare_RT", "RT"), ("Compare_LT", "LT"), ("Compare_HT", "HT")]:
                    prof_live, prof_blocks, prof_cols = build_compare_for_profile(
                        self._latest_type_to_frames, self._latest_components, self._latest_wo_map,
                        profile=prof, metrics=self._sel_metrics
                    )
                    render_ddmi_compare_sheet(
                        wb=writer.book,
                        sheet_name=sheet_name,
                        ddmi_blocks=prof_blocks,
                        components=self._latest_components,
                        active_metrics=self._sel_metrics,
                        active_sheets=prof_cols
                    )
            QMessageBox.information(self, "Complete", f"Compare Excel saved:\n{save_path}")
        except Exception as e:
            QMessageBox.critical(self, "Save Error", f"Failed to save:\n{e}")
        finally:
            self._end_busy()

# ───────────────────────────────────────────────────────────────
# Main Window with TWO TABS

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("All Tests — Summary & Compare")
        self.setMinimumSize(1280, 900)
        self.setStyleSheet(build_app_qss())

        tabs = QTabWidget()
        # 1st tab – existing Summary
        tabs.addTab(SummaryTab(), "All Summary & Data (Live)")
        # 2nd tab – existing Compare
        tabs.addTab(YieldFunnelWindow(), "Yield / Pareto / Schedule")
        # 3rd tab – trail0 (Yield + Pareto + Schedule)
        tabs.addTab(CompareTab(), "DDMI Columns Compare (Live)")
        # 4th tab (NO QSS here → preserves original Slot Efficiency sizing/UI)
        slot_tab = load_slot_efficiency_tab()
        tabs.addTab(slot_tab, "TP2TP3 Slot Efficiency")

        self.setCentralWidget(tabs)

# ───────────────────────────────────────────────────────────────
# Splash / Loading Screen

class SplashScreen(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
        self.setModal(True)
        self.setFixedSize(600, 320)
        self.setStyleSheet("background-color: #0f172a; color: white;")

        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 24, 24, 24)

        title = QLabel("All Tests — Summary & Compare")
        title.setStyleSheet("font-size: 22px; font-weight: 600;")
        subtitle = QLabel("Loading, please wait…")
        subtitle.setStyleSheet("font-size: 12px; color: #9ca3af;")

        lay.addWidget(title)
        lay.addWidget(subtitle)
        lay.addStretch(1)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setTextVisible(False)
        self.progress.setFixedHeight(10)
        self.progress.setStyleSheet("""
            QProgressBar {
                background: #1f2937;
                border-radius: 5px;
            }
            QProgressBar::chunk {
                background: #3b82f6;
                border-radius: 5px;
            }
        """)
        lay.addWidget(self.progress)

        author_row = QHBoxLayout()
        author_row.addStretch(1)
        author = QLabel("by Misbah Bilal")
        author.setStyleSheet("font-size: 11px; color: #9ca3af;")
        author_row.addWidget(author)
        lay.addLayout(author_row)

    def setProgress(self, val: int):
        self.progress.setValue(max(0, min(100, val)))
# ───────────────────────────────────────────────────────────────
# App bootstrap

if __name__ == "__main__":
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    app.setFont(QFont("Segoe UI", 10))

    # Splash screen
    splash = SplashScreen()
    # center splash on primary screen
    screen_geo = app.primaryScreen().availableGeometry()
    splash.move(
        screen_geo.center().x() - splash.width() // 2,
        screen_geo.center().y() - splash.height() // 2
    )
    splash.show()

    # smooth loading bar
    for i in range(0, 101, 4):
        splash.setProgress(i)
        QCoreApplication.processEvents()
        QThread.msleep(75)

    win = MainWindow()
    win.show()
    splash.close()

    sys.exit(app.exec_())