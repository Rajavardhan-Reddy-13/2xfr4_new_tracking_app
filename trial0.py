#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Analysis – Yield Funnel Dashboard + Fail Pareto + Schedule (trail0)

Features:
- Source mode:
    * "WO Lists": multiple named WO lists (per-list Yield + Pareto)
    * "Auto WOs (by device)": gets WOs from TESTRESULT_800G_MASTER
      by Device code (DR8+/FR4)
- Date range filter for Yield + Pareto using DDMI SID date (YYYYMMDD).
- Multi-station "Blank as PASS" override for selected stations.
- Fail Pareto:
    * Combined: FailMode@Station
    * Per-station: FailMode only for that station.
    * One fail mode per device (priority from your fail-mode list).
- Default fail modes are hard-coded; you can add/remove via UI.
- Right-click save:
    * tables → CSV
    * charts → PNG
- Schedule section:
    * Independent from main yield.
    * Single date selector.
    * Automatic WOs by device (DR8/FR4).
    * Table: Station | Cycle Time | Slots | Tested Qty.
    * Stations ordered in your real process flow.
"""
import os
import tempfile
import sys, re, traceback
from datetime import date, timedelta
import numpy as np
import pandas as pd
import pyodbc

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage

from PyQt5.QtCore import (
    Qt, QAbstractTableModel, QModelIndex, QSize, QThread,
    pyqtSignal, QDate
)
from PyQt5.QtGui import QFont, QBrush, QColor
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QTextEdit, QPushButton, QProgressBar, QTableView,
    QHeaderView, QMessageBox, QStyledItemDelegate, QStyleOptionViewItem,
    QTabWidget, QDialog, QDialogButtonBox, QLineEdit, QScrollArea,
    QMenu, QFileDialog, QComboBox, QCheckBox, QDateEdit, QGroupBox,
    QFormLayout,
    QTableWidget,
    QTableWidgetItem)
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from PyQt5.QtWidgets import QApplication, QAction, QAbstractItemView
from PyQt5.QtGui import QKeySequence


# ───────────────────────────────────────────────────────────────
# Paths & spec limits (shared with Summary GUI)

def get_app_folder() -> str:
    """Folder where the app/exe lives (used for local save files)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = get_app_folder()

# Local persistence for Yield tab (WO lists + Fail modes)
SAVED_LISTS_XLSX = os.path.join(APP_DIR, "yield_saved_lists.xlsx")


def _read_sheet_rows(ws):
    rows = []
    for r in ws.iter_rows(values_only=True):
        rows.append(list(r))
    return rows


def load_saved_wo_lists(path: str) -> list[dict]:
    """Load WO lists from Excel. Returns [{name:str, wos:[...]}]."""
    if not path or not os.path.exists(path):
        return []
    try:
        wb = load_workbook(path)
        if "WO_LISTS" not in wb.sheetnames:
            return []
        ws = wb["WO_LISTS"]
        rows = _read_sheet_rows(ws)
        if not rows or len(rows) < 2:
            return []

        header = [str(x).strip() if x is not None else "" for x in rows[0]]
        # expected columns: ListName, WO (others ignored)
        try:
            i_list = header.index("ListName")
            i_wo = header.index("WO")
        except ValueError:
            # backward compatibility: accept first two columns
            i_list, i_wo = 0, 1

        by_name: dict[str, list[str]] = {}
        for r in rows[1:]:
            if not r or len(r) <= max(i_list, i_wo):
                continue
            name = str(r[i_list]).strip() if r[i_list] is not None else ""
            wo = str(r[i_wo]).strip() if r[i_wo] is not None else ""
            if not name or not wo:
                continue
            by_name.setdefault(name, []).append(wo)

        out = []
        for name in by_name:
            # de-dupe while preserving order
            seen = set()
            wos = []
            for wo in by_name[name]:
                if wo in seen:
                    continue
                seen.add(wo)
                wos.append(wo)
            out.append({"name": name, "wos": wos})
        return out
    except Exception:
        return []


def load_saved_fail_modes(path: str) -> list[str]:
    """Load fail-mode keywords from Excel. Returns [FailMode, ...]."""
    if not path or not os.path.exists(path):
        return []
    try:
        wb = load_workbook(path)
        if "FAIL_MODES" not in wb.sheetnames:
            return []
        ws = wb["FAIL_MODES"]
        rows = _read_sheet_rows(ws)
        if not rows or len(rows) < 2:
            return []

        header = [str(x).strip() if x is not None else "" for x in rows[0]]
        try:
            i_mode = header.index("FailMode")
        except ValueError:
            i_mode = 0

        modes = []
        for r in rows[1:]:
            if not r or len(r) <= i_mode:
                continue
            m = str(r[i_mode]).strip() if r[i_mode] is not None else ""
            if m:
                modes.append(m)

        # de-dupe while preserving order
        seen = set()
        out = []
        for m in modes:
            mu = m.upper()
            if mu in seen:
                continue
            seen.add(mu)
            out.append(m)
        return out
    except Exception:
        return []


def load_saved_sn_lists(path: str) -> list[dict]:
    """Load SN lists from Excel. Returns [{name:str, sns:[...]}]."""
    if not path or not os.path.exists(path):
        return []
    try:
        wb = load_workbook(path)
        if "SN_LISTS" not in wb.sheetnames:
            return []
        ws = wb["SN_LISTS"]
        rows = _read_sheet_rows(ws)
        if not rows or len(rows) < 2:
            return []

        header = [str(x).strip() if x is not None else "" for x in rows[0]]
        try:
            i_list = header.index("ListName")
            i_sn = header.index("SN")
        except ValueError:
            i_list, i_sn = 0, 1

        by_name: dict[str, list[str]] = {}
        for r in rows[1:]:
            if not r or len(r) <= max(i_list, i_sn):
                continue
            name = str(r[i_list]).strip() if r[i_list] is not None else ""
            sn = str(r[i_sn]).strip() if r[i_sn] is not None else ""
            if not name or not sn:
                continue
            by_name.setdefault(name, []).append(sn)

        out = []
        for name in by_name:
            # de-dupe while preserving order
            seen = set()
            sns = []
            for sn in by_name[name]:
                if sn in seen:
                    continue
                seen.add(sn)
                sns.append(sn)
            out.append({"name": name, "sns": sns})
        return out
    except Exception:
        return []


def _open_or_create_wb(path: str):
    if os.path.exists(path):
        return load_workbook(path)
    wb = Workbook()
    # remove the default sheet
    try:
        wb.remove(wb.active)
    except Exception:
        pass
    return wb


def _replace_sheet(wb, name: str):
    if name in wb.sheetnames:
        ws_old = wb[name]
        wb.remove(ws_old)
    return wb.create_sheet(name)


def save_wo_lists(path: str, wo_lists: list[dict]):
    """Save WO lists into WO_LISTS sheet (preserves other sheets)."""
    wb = _open_or_create_wb(path)
    ws = _replace_sheet(wb, "WO_LISTS")
    ws.append(["ListName", "WOOrder", "WO"])
    for lst in wo_lists or []:
        name = str(lst.get("name", "")).strip()
        wos = lst.get("wos", []) or []
        if not name or not wos:
            continue
        for idx, wo in enumerate(wos, start=1):
            wo = str(wo).strip()
            if wo:
                ws.append([name, idx, wo])
    wb.save(path)


def save_fail_modes(path: str, fail_modes: list[str]):
    """Save fail modes into FAIL_MODES sheet (preserves other sheets)."""
    wb = _open_or_create_wb(path)
    ws = _replace_sheet(wb, "FAIL_MODES")
    ws.append(["Order", "FailMode"])
    for idx, m in enumerate(fail_modes or [], start=1):
        m = str(m).strip()
        if m:
            ws.append([idx, m])
    wb.save(path)

def save_sn_lists(path: str, sn_lists: list[dict]):
    """Save SN lists into SN_LISTS sheet (preserves other sheets)."""
    wb = _open_or_create_wb(path)
    ws = _replace_sheet(wb, "SN_LISTS")
    ws.append(["ListName", "SNOrder", "SN"])
    for lst in sn_lists or []:
        name = str(lst.get("name", "")).strip()
        sns = lst.get("sns", []) or []
        if not name or not sns:
            continue
        for idx, sn in enumerate(sns, start=1):
            sn = str(sn).strip()
            if sn:
                ws.append([name, idx, sn])
    wb.save(path)

# If your Summary GUI already uses a spec txt, just make sure the
# filename below matches that file (same folder as this script).
SPEC_FILE_NAME = "spec_limits.txt"
SPEC_FILE = os.path.join(APP_DIR, SPEC_FILE_NAME)

# (station_lower, metric_lower) -> (lsl, usl)
SPEC_LIMITS: dict[tuple[str, str], tuple[float | None, float | None]] = {}


def _to_float_or_none(x: str):
    x = (x or "").strip()
    if not x or x.upper() in ("NA", "N/A", "NONE", "NULL"):
        return None
    try:
        return float(x)
    except Exception:
        return None


def load_spec_limits():
    """
    Load spec limits from SPEC_FILE into SPEC_LIMITS.

    Supported formats (one line per metric):

        Metric, LSL, USL
    or
        Station, Metric, LSL, USL

    Separator can be comma, semicolon, or tab.
    Lines starting with '#' are ignored.
    """
    global SPEC_LIMITS
    SPEC_LIMITS = {}

    if not os.path.exists(SPEC_FILE):
        return

    try:
        with open(SPEC_FILE, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue

                parts = re.split(r"[,\t;]+", line)
                if len(parts) == 3:
                    station = "*"
                    metric, lsl_s, usl_s = parts
                elif len(parts) >= 4:
                    station, metric, lsl_s, usl_s = parts[:4]
                else:
                    continue

                metric = metric.strip()
                station = (station or "*").strip()
                if not metric:
                    continue

                lsl = _to_float_or_none(lsl_s)
                usl = _to_float_or_none(usl_s)
                key = (station.lower(), metric.lower())
                SPEC_LIMITS[key] = (lsl, usl)
    except Exception:
        # Don't kill the app if the file is malformed
        SPEC_LIMITS = {}


def get_spec_for_metric(metric_name: str, station_name: str | None = None):
    """
    Return (lsl, usl) for metric/station, or None if not found.

    Lookup order:
        (station, metric)  -> ('*', metric) -> ('', metric)
    """
    if not SPEC_LIMITS or not metric_name:
        return None

    m = metric_name.strip().lower()
    if not m:
        return None

    station_candidates = []
    if station_name:
        station_candidates.append(station_name.strip().lower())
    station_candidates.extend(["*", ""])

    for st in station_candidates:
        key = (st, m)
        if key in SPEC_LIMITS:
            return SPEC_LIMITS[key]

    return None


# Load specs once at startup
load_spec_limits()


# ───────────────────────────────────────────────────────────────
# DB connection and source tables

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
BURNIN_TABLE2  = "TestResult_800G_BURNIN_TEST"       # legacy burn-in table (also used)
SWITCH_TABLE2  = "TestResult_800G_TRX_Switch_TEST"   # legacy switch test table (also used)
MASTER_WO_TABLE = "TESTRESULT_800G_MASTER"  # COMPONENTID, WO, Device

# Device IDs for AUTO-WO mode & Schedule (edit if needed)
DR8_DEVICE_ID = "400454000023"
FR4_DEVICE_ID = "400454000024"
SCHEDULE_DEVICE_IDS = [DR8_DEVICE_ID, FR4_DEVICE_ID]

SUMMARY_COLUMNS = [
    ("FW Writing",   ("FWWRITE", None)),
    ("DDMI Cal",     ("TRX", "DDMI")),
    ("TP2TP3 - RT",  ("TRX", "RT")),
    ("TP2TP3 - LT",  ("TRX", "LT")),
    ("TP2TP3 - HT",  ("TRX", "HT")),
    ("Burn-in",      ("BURNIN", None)),
    ("3TBER",        ("3TBER", None)),
    ("TCBER",        ("TCBER", None)),
    ("Mode Hopping", ("MODEHOP", None)),
    ("Final Test",   ("TRX", "FINAL")),
    ("Switch Test",  ("SWITCH", None)),
]

# For Yield funnel order / labels and schedule order
STATION_FLOW = [
    ("FW Writing",   "FW Writing"),
    ("DDMI Cal",     "DDMI Cal"),
    ("TP2TP3 RT",    "TP2TP3 - RT"),
    ("TP2TP3 LT",    "TP2TP3 - LT"),
    ("TP2TP3 HT",    "TP2TP3 - HT"),
    ("Burn-in",      "Burn-in"),
    ("3T Test",      "3TBER"),
    ("TCBER",        "TCBER"),
    ("Mode Hopping", "Mode Hopping"),
    ("Final Test",   "Final Test"),
    ("Switch Test",  "Switch Test"),
]

# Map station labels <-> SUMMARY column names
LABEL_TO_COL = dict(STATION_FLOW)                       # "TP2TP3 LT"  -> "TP2TP3 - LT"
COL_TO_LABEL = {col: label for (label, col) in STATION_FLOW}  # "TP2TP3 - LT" -> "TP2TP3 LT"

# List of SUMMARY status column names (for colouring in detail tables)
SUMMARY_STATUS_COLUMNS = [name for (name, _sub) in SUMMARY_COLUMNS]


# DEFAULT FAIL MODES (keywords, case-insensitive)
# You can edit this list; UI lets you adjust later too.
DEFAULT_FAIL_MODES = [
    "TCOrder",
    "TDECQ",
    "Sen",
    "ModeHopping",
    "RLM",
    "DRXP",
    "LOS",
    "VCC_offset",
    "Tx_Power",
]

# For Schedule table – set real values here
STATION_SLOTS = {
    "DDMI Cal":     2,
    "FW Writing":   3,
    "TP2TP3 RT":    2,
    "TP2TP3 LT":    2,
    "TP2TP3 HT":    2,
    "Mode Hopping": 1,
    "Burn-in":      144,
    "3T Test":      8,
    "TCBER":        1,
    "Final Test":   2,
    "Switch Test":  13,
}
STATION_CYCLE = {
    "DDMI Cal":     "10 min/device",
    "FW Writing":   "3 min/device",
    "TP2TP3 RT":    "1 hr/3 devices",
    "TP2TP3 LT":    "1 hr/3 devices",
    "TP2TP3 HT":    "1 hr/3 devices",
    "Burn-in":      "12 hr/device",
    "3T Test":      "15 min (4 dev)",
    "TCBER":        "1.5 hr/5 devices",
    "Mode Hopping": "1 dev/15 min",
    "Final Test":   "1 hr/3 devices",
    "Switch Test":  "2 hr",
}

# patterns
SID_NUM_PAT   = re.compile(r"(\d{8,})")
SID_DATE_PAT  = re.compile(r"^[A-Za-z](20\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01]))")
STAGE_PAT     = re.compile(r"^\s*(\d+)\s*_(RT|LT|HT|ATS)\s*$", re.IGNORECASE)

BURNIN_REQUIRED_CYCLES = set(range(24))
TXP_MIN, TXP_MAX = -2.4, 4.0
RXP_MIN, RXP_MAX = -2.4, 4.0

# ───────────────────────────────────────────────────────────────
# helpers

def build_station_detail_df(type_to_frames, sn: str, station_name: str) -> pd.DataFrame:
    """
    Return the same raw dataframe that Summary uses for the station detail
    (including DDMI gate logic). This MUST match the logic in SummaryTab.
    """
    sn = str(sn).strip()
    if not sn:
        return pd.DataFrame()

    # Same mapping as in SummaryTab
    col_map = dict(SUMMARY_COLUMNS)          # e.g. "Burn-in" -> ("BURNIN", None)
    family, subtype = col_map[station_name]

    if family == "TRX":
        trx_map = type_to_frames.get("TRX", {})
        df_all = trx_map.get(subtype, pd.DataFrame())
    else:
        df_all = type_to_frames.get(family, pd.DataFrame())

    if df_all is None or df_all.empty:
        return pd.DataFrame()

    sub = df_all[df_all["COMPONENTID"].astype(str).str.strip() == sn].copy()
    if sub.empty:
        return sub

    # ---------- DDMI gate, same as Summary ----------
    try:
        trx_map = type_to_frames.get("TRX", {})
        gate_map = compute_ddmi_gate_sid_for_components(trx_map, [sn])
        gate_sid = gate_map.get(sn)
    except Exception:
        gate_sid = None

    if gate_sid is not None and "_SID_NUM_" in sub.columns and station_name not in ("DDMI Cal", "FW Writing", "Mode Hopping"):
        sid_num = pd.to_numeric(sub["_SID_NUM_"], errors="coerce")
        sub = sub[sid_num >= gate_sid].copy()

    return sub

def sid_to_number(sid):
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

def sid_to_yyyymmdd_int(sid):
    if sid is None:
        return None
    s = str(sid)
    m = SID_DATE_PAT.search(s)
    if not m:
        return None
    try:
        return int(m.group(1))
    except Exception:
        return None

def sid_to_ymd_time(sid):
    """
    Extract (YYYYMMDD, minutes_since_midnight) from SID if possible.
    Time is parsed from digits after YYYYMMDD, using first 4–6 digits as HHMMxx.
    """
    if sid is None:
        return None, None
    s = str(sid)
    m = SID_DATE_PAT.search(s)
    if not m:
        return None, None
    date_str = m.group(1)
    end = m.end()
    time_digits = ""
    for ch in s[end:]:
        if ch.isdigit():
            time_digits += ch
            if len(time_digits) >= 6:
                break
        else:
            break
    ymd = int(date_str)
    if len(time_digits) < 4:
        return ymd, None
    try:
        hh = int(time_digits[0:2])
        mm = int(time_digits[2:4])
    except Exception:
        return ymd, None
    if not (0 <= hh <= 23 and 0 <= mm <= 59):
        return ymd, None
    minutes = hh * 60 + mm
    return ymd, minutes

def parse_lane_stage(ch: str):
    if ch is None:
        return (None, None)
    s = str(ch).strip().upper()
    m = STAGE_PAT.match(s)
    if m:
        lane = int(m.group(1))
        stage = m.group(2).upper()
        return (lane if 1 <= lane <= 8 else None, stage)
    m2 = re.search(r"(?<!\d)([1-8])(?!\d)", s)
    lane = int(m2.group(1)) if m2 else None
    return (lane, None)

def _find_ci_col(df: pd.DataFrame, targets):
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
    if "FailureCodeID" in out.columns:
        out["FailureCodeID"] = (
            out["FailureCodeID"]
            .apply(lambda x: "" if pd.isna(x) else str(x).strip())
            .replace({"nan": "", "NaN": "", "None": "", "NONE": "", "NULL": "", "null": ""})
        )
    else:
        out["FailureCodeID"] = ""
    if "CHNumber" in out.columns:
        out["_CH_IDX_"] = out["CHNumber"].apply(lambda x: parse_lane_stage(x)[0])
    else:
        out["_CH_IDX_"] = None
    if "SID" in out.columns:
        out["_SID_NUM_"] = out["SID"].apply(sid_to_number)
    else:
        out["_SID_NUM_"] = pd.NA
    return out

def keep_latest_per_channel(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    key_cols = [c for c in ["COMPONENTID", "CHNumber"] if c in df.columns]
    if not key_cols:
        return normalize_test_df(df)
    if "SID" in df.columns:
        sidnum = df["SID"].apply(sid_to_number)
        df2 = (
            df.assign(_SID_=sidnum)
              .sort_values(key_cols + ["_SID_"])
              .drop_duplicates(subset=key_cols, keep="last")
              .drop(columns=["_SID_"])
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
    # normalize column name (some tables use ComponentID)
    if "COMPONENTID" not in d.columns:
        for c in d.columns:
            if c.lower() == "componentid":
                d = d.rename(columns={c: "COMPONENTID"})
                break
    if "COMPONENTID" not in d.columns:
        return d

    # Prefer Create_Time, else SID numeric, else original order
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
    """Legacy helper (kept for backwards-compat).

    If ModeHopping TESTNUMBER is non-numeric (often SID-like), filtering by max(TESTNUMBER)
    can incorrectly return empty. This safe version keeps the latest row per COMPONENTID based
    on SID ordering (or Timestamp if SID missing).
    """
    if df is None or df.empty:
        return df
    if 'COMPONENTID' not in df.columns:
        return df
    # Prefer SID ordering
    if 'SID' in df.columns:
        tmp=df.copy()
        tmp['_sid_num']=tmp['SID'].apply(sid_to_number)
        tmp=tmp.sort_values(['COMPONENTID','_sid_num'], ascending=[True,True])
        tmp=tmp.dropna(subset=['_sid_num']) if tmp['_sid_num'].notna().any() else tmp
        tmp=tmp.groupby('COMPONENTID', as_index=False).tail(1)
        return tmp.drop(columns=['_sid_num'], errors='ignore')
    # Fallback: Timestamp
    ts_col=None
    for c in ['Timestamp','Create_Time','CreateTime','timestamp']:
        if c in df.columns:
            ts_col=c;break
    if ts_col:
        tmp=df.copy()
        tmp['_ts']=pd.to_datetime(tmp[ts_col], errors='coerce')
        tmp=tmp.sort_values(['COMPONENTID','_ts'], ascending=[True,True])
        tmp=tmp.groupby('COMPONENTID', as_index=False).tail(1)
        return tmp.drop(columns=['_ts'], errors='ignore')
    return df


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
    # normalize ComponentID column
    comp_col = "COMPONENTID" if "COMPONENTID" in d.columns else next((c for c in d.columns if c.lower()=="componentid"), None)
    if comp_col is None:
        return ""
    sub = d[d[comp_col].astype(str).str.strip() == cid]
    if sub.empty:
        return ""
    # If table has WO column, enforce match
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

def split_trx_types(df_trx: pd.DataFrame) -> dict:
    # NOTE:
    # TRX table contains multiple logical stages. Historically we split by CHNumber suffix
    # (_DDMI/_RT/_LT/_HT/_ATS). For FW Writing + Mode Hopping, those suffixes may not exist,
    # so we also try to split using Station/TestName/Type columns if present.
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

    # Build a raw "source" string from any available columns for extra matching.
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

        # ----- FINAL TEST SID RULE (Yield page, aggregated) -----
        if key == "FINAL" and "SID" in sub.columns:
            sid_str = sub["SID"].astype(str).str.strip()
            base_sid = sid_str.str.replace(r"A$", "", regex=True)
            ends_A = sid_str.str.endswith("A")

            tmp = pd.DataFrame({"base": base_sid, "ends_A": ends_A})
            has_nonA = tmp.groupby("base")["ends_A"].transform(lambda s: (~s).any())
            mask = (has_nonA & ~ends_A) | (~has_nonA)
            sub = sub[mask]

        result[key] = keep_latest_per_channel(sub) if not sub.empty else pd.DataFrame()

    return result


def split_trx_types_raw(df_trx: pd.DataFrame) -> dict:
    result = {k: pd.DataFrame() for k in ["DDMI", "RT", "LT", "HT", "FINAL", "MODEHOP", "FW"]}
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

    if type_col:
        tser = df[type_col].astype(str).str.upper()
    else:
        tser = pd.Series(index=df.index, dtype=object)

    tser[ch.str.endswith("_DDMI")] = "DDMI"
    tser[ch.str.endswith("_RT")]   = "RT"
    tser[ch.str.endswith("_LT")]   = "LT"
    tser[ch.str.endswith("_HT")]   = "HT"
    tser[ch.str.contains("_ATS", na=False)] = "FINAL"

    stage_keys = {"DDMI", "RT", "LT", "HT", "FINAL"}
    can_set = ~tser.fillna("").isin(list(stage_keys))

    mh_mask = raw_src.str.contains("MODE", na=False) & raw_src.str.contains("HOP", na=False)
    fw_mask = raw_src.str.contains("FW", na=False) & (raw_src.str.contains("WRIT", na=False) | raw_src.str.contains("FIRM", na=False))

    tser[mh_mask & can_set] = "MODEHOP"
    tser[fw_mask & can_set] = "FW"

    df = df.assign(_Type_=tser.fillna("TRX"))

    for key in ["DDMI", "RT", "LT", "HT", "FINAL", "MODEHOP", "FW"]:
        sub = df[df["_Type_"] == key].copy()

        # ----- FINAL TEST SID RULE (Yield page, raw views) -----
        if key == "FINAL" and "SID" in sub.columns:
            sid_str = sub["SID"].astype(str).str.strip()
            base_sid = sid_str.str.replace(r"A$", "", regex=True)
            ends_A = sid_str.str.endswith("A")

            tmp = pd.DataFrame({"base": base_sid, "ends_A": ends_A})
            has_nonA = tmp.groupby("base")["ends_A"].transform(lambda s: (~s).any())
            mask = (has_nonA & ~ends_A) | (~has_nonA)
            sub = sub[mask]

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

def compute_ddmi_date_for_components(trx_map: dict, components: list) -> dict:
    dd = trx_map.get("DDMI", pd.DataFrame())
    out = {}
    if dd is None or dd.empty:
        return {str(c).strip(): None for c in components}
    tmp = dd.copy()
    tmp["_SID_DATE_"] = tmp["SID"].apply(sid_to_yyyymmdd_int)
    tmp["COMPONENTID"] = tmp["COMPONENTID"].astype(str).str.strip()
    for comp, grp in tmp.groupby("COMPONENTID"):
        dates = pd.to_numeric(grp["_SID_DATE_"], errors="coerce").dropna()
        out[comp] = int(dates.min()) if not dates.empty else None
    for c in components:
        key = str(c).strip()
        out.setdefault(key, None)
    return out

# ───────────────────────────────────────────────────────────────
# Summary helpers

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
        codes = [c for c in codes if c not in ("", "NAN", "NaN", "None", "NONE", "NULL", "null")]
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
    if "CHNumber" in df.columns:
        parsed = df["CHNumber"].apply(parse_lane_stage)
        df["_CH_IDX_"] = [p[0] for p in parsed]
        df["_STAGE_"]  = [p[1] for p in parsed]
    else:
        df["_STAGE_"] = None
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
        codes = [c for c in codes if c not in ("", "NAN", "NaN", "None", "NONE", "NULL", "null")]
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
    """
    Burn-in summary:
      - DDMI gate SID is IGNORED
      - Still require 8 lanes and full 0..23 cycles on each lane
      - PASS if all those rows are PASS, else list worst TXP/RXP per lane
    """
    if df is None or df.empty:
        return ""

    df = _rename_burnin_channel(df)
    df = normalize_test_df(df).copy()

    cycle_col = _find_ci_col(df, ["Test cycle", "Test_cycle", "Cycle", "TEST CYCLE"])
    if cycle_col is None:
        return ""

    # ⚠️ ddmi_gate_sid is intentionally ignored here – no SID filtering

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

    if (status == "PASS").all():
        return "PASS"

    tx_col = _find_ci_col(sub, ["DDMI_TxP(dbm)", "DDMI_TxP(dBm)", "TxP", "TxPower"])
    rx_col = _find_ci_col(sub, ["DDMI_RxP(dbm)", "DDMI_RxP(dBm)", "RxP", "RxPower"])
    out_items, seen = [], set()

    fail_rows = sub[status != "PASS"].copy()
    for lane, grp in fail_rows.groupby("_CH_IDX_"):
        lane_tag = f"CH{int(lane)}"
        if tx_col:
            worst_tx, cyc_tx = _worst_oob_with_cycle(grp, tx_col, cycle_col, TXP_MIN, TXP_MAX)
            if worst_tx is not None:
                item = f"{lane_tag} DDMI TXP {worst_tx:.2f}"
                if cyc_tx is not None:
                    item += f" (cycle {cyc_tx})"
                if item not in seen:
                    seen.add(item)
                    out_items.append(item)
        if rx_col:
            worst_rx, cyc_rx = _worst_oob_with_cycle(grp, rx_col, cycle_col, RXP_MIN, RXP_MAX)
            if worst_rx is not None:
                item = f"{lane_tag} DDMI RXP {worst_rx:.2f}"
                if cyc_rx is not None:
                    item += f" (cycle {cyc_rx})"
                if item not in seen:
                    seen.add(item)
                    out_items.append(item)

    return ", ".join(out_items) if out_items else ""


def summarize_final_ats(df: pd.DataFrame, ddmi_gate_sid: float | None) -> str:
    if df is None or df.empty:
        return ""
    df = normalize_test_df(df).copy()
    ch = df["CHNumber"].astype(str).str.upper()
    mask = (
        ch.str.endswith("_ATS")
        | ch.str.endswith("_ATS_RT")
        | ch.str.endswith("_ATS_LT")
        | ch.str.endswith("_ATS_HT")
    )
    df = df[mask]
    if df.empty:
        return ""

    # Apply gate SID filter only when gate is available and produces results
    if ddmi_gate_sid is not None:
        gated = df[pd.to_numeric(df["_SID_NUM_"], errors="coerce") >= ddmi_gate_sid]
        if not gated.empty:
            df = gated

    df["_SID_NUM_"] = pd.to_numeric(df["_SID_NUM_"], errors="coerce")
    df = df.sort_values("_SID_NUM_", ascending=False)
    latest = df.head(24)

    status = latest["CH_Pass_Fail"].astype(str).str.upper().str.strip()
    has_fail = (status != "PASS").any()

    # Only declare PASS when all 24 channels present and all pass
    if not has_fail:
        return "PASS" if len(latest) >= 24 else ""

    # Has fails — report them regardless of row count or gate SID
    fail_df = latest[status != "PASS"]
    msgs = []
    for _, row in fail_df.iterrows():
        chn = row["CHNumber"]
        code = row["FailureCodeID"]
        code = "" if code in ("", None, "NULL", "nan", "NaN") else str(code)
        msgs.append(f"{chn} {code}" if code else str(chn))
    return ", ".join(msgs)

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

# ───────────────────────────────────────────────────────────────
# DB helpers

def _chunk(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i+n]

def fetch_components_by_wos(conn, wos: list[str]) -> list[str]:
    if not wos:
        return []
    wos = [w.strip() for w in wos if w.strip()]
    if not wos:
        return []
    cursor = conn.cursor()
    found = set()
    for group in _chunk(wos, 800):
        placeholders = ",".join(["?"] * len(group))
        sql = (
            f"SELECT DISTINCT RTRIM(LTRIM(COMPONENTID)) "
            f"FROM {MASTER_WO_TABLE} "
            f"WHERE LTRIM(RTRIM(WO)) IN ({placeholders})"
        )
        cursor.execute(sql, group)
        rows = cursor.fetchall()
        for (comp,) in rows:
            if comp is not None:
                found.add(str(comp).strip())
    return sorted(found)


def fetch_components_by_devices(conn, device_ids: list[str]):
    if not device_ids:
        return [], {}
    cursor = conn.cursor()
    placeholders = ",".join(["?"] * len(device_ids))
    sql = (
        f"SELECT DISTINCT "
        f"LTRIM(RTRIM(COMPONENTID)) AS COMPONENTID, "
        f"LTRIM(RTRIM(WO))          AS WO "
        f"FROM {MASTER_WO_TABLE} "
        f"WHERE Device IN ({placeholders})"
    )
    cursor.execute(sql, [d.strip() for d in device_ids])
    rows = cursor.fetchall()
    comps = []
    wo_map: dict[str, str] = {}
    for comp, wo in rows:
        comp_str = "" if comp is None else str(comp).strip()
        if not comp_str:
            continue
        if comp_str not in wo_map:
            comps.append(comp_str)
        wo_map[comp_str] = "" if wo is None else str(wo).strip()
    comps = sorted(comps)
    return comps, wo_map



def fetch_wos_by_devices(conn, device_ids: list[str]) -> list[str]:
    """Return distinct WO strings for the given Device IDs using the same master table as Auto-WO."""
    if not device_ids:
        return []
    cursor = conn.cursor()
    placeholders = ",".join(["?"] * len(device_ids))
    sql = (
        f"SELECT DISTINCT LTRIM(RTRIM(WO)) AS WO "
        f"FROM {MASTER_WO_TABLE} "
        f"WHERE Device IN ({placeholders})"
    )
    cursor.execute(sql, [d.strip() for d in device_ids if str(d).strip()])
    rows = cursor.fetchall()
    wos = []
    seen = set()
    for (wo,) in rows:
        s = "" if wo is None else str(wo).strip()
        if not s:
            continue
        if s not in seen:
            seen.add(s)
            wos.append(s)
    return sorted(wos)


def fetch_table_for_components_bulk(conn, table: str, components: list[str], chunk_size: int = 800) -> pd.DataFrame:
    if not components:
        return pd.DataFrame()
    cursor = conn.cursor()
    out_frames = []
    comps = [str(x).strip() for x in components if str(x).strip()]
    for group in _chunk(comps, chunk_size):
        placeholders = ",".join(["?"] * len(group))
        sql = f"SELECT * FROM {table} WHERE RTRIM(LTRIM(COMPONENTID)) IN ({placeholders})"
        cursor.execute(sql, group)
        rs = cursor.fetchall()
        if rs:
            df = pd.DataFrame.from_records(rs, columns=[c[0] for c in cursor.description])
            out_frames.append(df)
    return pd.concat(out_frames, ignore_index=True) if out_frames else pd.DataFrame()

def fetch_burnin_combined(conn, components: list[str]) -> pd.DataFrame:
    """Fetch burn-in rows from both tables and combine, deduplicating by (COMPONENTID, SID)."""
    df1 = fetch_table_for_components_bulk(conn, TABLES["BURNIN"], components)
    df2 = fetch_table_for_components_bulk(conn, BURNIN_TABLE2, components)
    if df1.empty and df2.empty:
        return pd.DataFrame()
    combined = pd.concat([df1, df2], ignore_index=True)
    if "COMPONENTID" in combined.columns and "SID" in combined.columns:
        combined = combined.drop_duplicates(subset=["COMPONENTID", "SID"])
    return combined.reset_index(drop=True)

def fetch_switch_combined(conn, components: list[str]) -> pd.DataFrame:
    """Fetch switch test rows from TestResult_800G_2XFR4_TRX_Switch_TEST."""
    return fetch_table_for_components_bulk(conn, TABLES["SWITCH"], components)

def fetch_table_for_date(conn, table: str, target_ymd: int) -> pd.DataFrame:
    """Fetch all rows from a table where the SID encodes target_ymd (format A{YYYYMMDD}...).
    Uses SUBSTRING(SID, 2, 8) = date_str so no master-table lookup is needed."""
    date_str = str(target_ymd)
    try:
        cursor = conn.cursor()
        sql = (
            f"SELECT * FROM {table} "
            f"WHERE LEN(ISNULL(SID,'')) >= 9 AND SUBSTRING(SID, 2, 8) = ?"
        )
        cursor.execute(sql, [date_str])
        cols = [d[0] for d in cursor.description]
        rows = cursor.fetchall()
        if not rows:
            return pd.DataFrame(columns=cols)
        return pd.DataFrame.from_records(rows, columns=cols)
    except Exception:
        return pd.DataFrame()


def fetch_burnin_for_date(conn, target_ymd: int) -> pd.DataFrame:
    df1 = fetch_table_for_date(conn, TABLES["BURNIN"], target_ymd)
    df2 = fetch_table_for_date(conn, BURNIN_TABLE2,    target_ymd)
    if df1.empty and df2.empty:
        return pd.DataFrame()
    combined = pd.concat([df1, df2], ignore_index=True)
    if "COMPONENTID" in combined.columns and "SID" in combined.columns:
        combined = combined.drop_duplicates(subset=["COMPONENTID", "SID"])
    return combined.reset_index(drop=True)


def fetch_switch_for_date(conn, target_ymd: int) -> pd.DataFrame:
    return fetch_table_for_date(conn, TABLES["SWITCH"], target_ymd)


def fetch_wo_map(conn, components: list[str], restrict_wos=None) -> dict[str, str]:
    if not components:
        return {}

    # Unique, trimmed SN list
    comps_unique, seen = [], set()
    for c in components:
        s = str(c).strip()
        if s and s not in seen:
            comps_unique.append(s)
            seen.add(s)

    wo_map: dict[str, str] = {}
    cursor = conn.cursor()

    # Optional WO filter (for WO-mode)
    restrict_wos = [w.strip() for w in (restrict_wos or []) if w.strip()]
    use_wo_filter = bool(restrict_wos)

    for group in _chunk(comps_unique, 800):
        comp_placeholders = ",".join(["?"] * len(group))

        if use_wo_filter:
            wo_placeholders = ",".join(["?"] * len(restrict_wos))
            sql = (
                f"SELECT LTRIM(RTRIM(COMPONENTID)) AS COMPONENTID, "
                f"       LTRIM(RTRIM(WO))          AS WO "
                f"FROM {MASTER_WO_TABLE} "
                f"WHERE LTRIM(RTRIM(COMPONENTID)) IN ({comp_placeholders}) "
                f"  AND LTRIM(RTRIM(WO)) IN ({wo_placeholders})"
            )
            params = [s.strip() for s in group] + restrict_wos
        else:
            sql = (
                f"SELECT LTRIM(RTRIM(COMPONENTID)) AS COMPONENTID, "
                f"       LTRIM(RTRIM(WO))          AS WO "
                f"FROM {MASTER_WO_TABLE} "
                f"WHERE LTRIM(RTRIM(COMPONENTID)) IN ({comp_placeholders})"
            )
            params = [s.strip() for s in group]

        cursor.execute(sql, params)
        rows = cursor.fetchall()
        for comp, wo in rows:
            comp_s = str(comp).strip()
            wo_s = "" if wo is None else str(wo).strip()
            # keep first seen mapping; or change this if you prefer “latest”
            if comp_s not in wo_map:
                wo_map[comp_s] = wo_s

    # Ensure every SN has a key (possibly empty)
    for c in comps_unique:
        wo_map.setdefault(c, "")

    return wo_map

# ───────────────────────────────────────────────────────────────
# Yield funnel

def compute_station_masks(summary_df: pd.DataFrame,
                          blank_pass_labels=None,
                          bypass_labels=None,
                          switch_test_pass=False) -> dict[str, dict[str, pd.Series]]:
    """
    For each station label in STATION_FLOW, compute boolean masks:
        {
          label: {
             "input": input_mask,
             "wip":   wip_mask,
             "pass":  pass_mask,
             "fail":  fail_mask,
          }
        }
    This mirrors the logic in build_yield_funnel.
    """
    if summary_df is None or summary_df.empty:
        return {}

    if blank_pass_labels is None:
        blank_pass_labels = set()
    else:
        blank_pass_labels = set(blank_pass_labels)

    if bypass_labels is None:
        bypass_labels = set()
    else:
        bypass_labels = set(bypass_labels)

    df = summary_df.copy()
    station_info: dict[str, dict[str, pd.Series]] = {}

    switch_test_pass_mask = pd.Series(False, index=df.index)
    if switch_test_pass and "Switch Test" in df.columns:
        switch_test_pass_mask = (df["Switch Test"].str.strip().str.upper() == "PASS").fillna(False)

    active_mask = pd.Series(True, index=df.index, dtype=bool)

    for label, col_name in STATION_FLOW:
        input_mask = active_mask.copy()

        # Bypass station: Input = Output, Fail = 0, no gating
        if label in bypass_labels:
            pass_here = input_mask
            wip_here = pd.Series(False, index=df.index)
            fail_here = pd.Series(False, index=df.index)

        elif col_name not in df.columns:
            wip_here = input_mask
            pass_here = pd.Series(False, index=df.index)
            fail_here = input_mask
        else:
            col = df[col_name].fillna("").astype(str)
            col_str = col.str.strip()
            col_up = col_str.str.upper()

            if label in blank_pass_labels:
                # blanks treated as PASS
                is_blank = (col_str == "")
                pass_here = input_mask & ((col_up == "PASS") | is_blank)
                wip_here = pd.Series(False, index=df.index)
                fail_here = input_mask & ~pass_here
            else:
                wip_here = input_mask & (col_str == "")
                pass_here = input_mask & (col_up == "PASS")
                fail_here = input_mask & ~(wip_here | pass_here)

        # Override for switch test pass
        if switch_test_pass:
            pass_here = pass_here | switch_test_pass_mask
            wip_here = wip_here & ~switch_test_pass_mask
            fail_here = fail_here & ~switch_test_pass_mask

        station_info[label] = {
            "input": input_mask,
            "wip":   wip_here,
            "pass":  pass_here,
            "fail":  fail_here,
        }

        # Only PASS devices continue to next station
        active_mask = pass_here
        if switch_test_pass:
            active_mask = active_mask | switch_test_pass_mask

    return station_info


def build_yield_funnel(summary_df: pd.DataFrame,
                       wo_subset=None,
                       blank_pass_labels=None,
                       bypass_labels=None,
                       switch_test_pass=False) -> pd.DataFrame:
    if summary_df is None or summary_df.empty:
        return pd.DataFrame(columns=["Process","Station","Input Qty","WIP","Output Qty","Fail Qty","Yield"])

    # Normalise blank-pass station list
    if blank_pass_labels is None:
        blank_pass_labels = set()
    else:
        blank_pass_labels = set(blank_pass_labels)

    # Normalise bypass station list
    if bypass_labels is None:
        bypass_labels = set()
    else:
        bypass_labels = set(bypass_labels)

    df = summary_df.copy()

    # Optional WO subset
    if wo_subset:
        if isinstance(wo_subset, str):
            wo_subset = [wo_subset]
        keep = df["WO"].astype(str).str.strip().isin([w.strip() for w in wo_subset])
        df = df[keep].reset_index(drop=True)

    if df.empty:
        return pd.DataFrame(columns=["Process","Station","Input Qty","WIP","Output Qty","Fail Qty","Yield"])

    n_devices = len(df)

    station_info = compute_station_masks(df, blank_pass_labels, bypass_labels, switch_test_pass)

    rows = []
    fail_any = pd.Series(False, index=df.index)
    wip_any = pd.Series(False, index=df.index)

    for label, _col_name in STATION_FLOW:
        info = station_info.get(label)
        if info is None:
            # Should not happen, but keep safe
            rows.append({
                "Process": "Testing",
                "Station": label,
                "Input Qty": 0,
                "WIP": 0,
                "Output Qty": 0,
                "Fail Qty": 0,
                "Yield": "",
            })
            continue

        input_mask = info["input"]
        wip_mask   = info["wip"]
        pass_mask  = info["pass"]
        fail_mask  = info["fail"]

        input_qty  = int(input_mask.sum())
        wip_qty    = int(wip_mask.sum())
        fail_qty   = int(fail_mask.sum())
        output_qty = int(pass_mask.sum())

        fail_any |= fail_mask
        wip_any  |= wip_mask

        if input_qty == 0:
            yield_str = ""
        else:
            denom = output_qty + fail_qty
            yield_str = f"{(output_qty/denom*100):.2f}%" if denom > 0 else ""

        rows.append({
            "Process": "Testing",
            "Station": label,
            "Input Qty": input_qty,
            "WIP": wip_qty,
            "Output Qty": output_qty,
            "Fail Qty": fail_qty,
            "Yield": yield_str,
        })

    # Total row
    total_input  = n_devices
    total_output = int(station_info.get(STATION_FLOW[-1][0], {}).get("pass", pd.Series(False, index=df.index)).sum())
    total_fail   = int(fail_any.sum())
    total_wip    = int(wip_any.sum())
    denom_tot    = total_output + total_fail
    total_yield  = f"{(total_output/denom_tot*100):.2f}%" if denom_tot > 0 else ""

    rows.append({
        "Process": "",
        "Station": "Total",
        "Input Qty": total_input,
        "WIP": total_wip,
        "Output Qty": total_output,
        "Fail Qty": total_fail,
        "Yield": total_yield,
    })

    return pd.DataFrame(rows,
                        columns=["Process","Station","Input Qty","WIP","Output Qty","Fail Qty","Yield"])


def compute_fpy_table(summary_df: pd.DataFrame,
                      wo_subset=None,
                      blank_pass_labels=None,
                      bypass_labels=None,
                      ever_failed_by_station: dict = None) -> pd.DataFrame:
    """Per-station True First Pass Yield plus an Overall row.

    ever_failed_by_station: {station_label: set of COMPONENTIDs that had ANY fail
    row in the raw (all-attempts) data}.  When provided, a device is first-pass at
    a station only if it was tested there AND never appeared in that station's fail set.
    This correctly excludes devices that failed on attempt 1 and passed on attempt 2.
    """
    cols_out = ["Station", "Tested", "First Pass", "Re-test Pass/Fail", "FPY"]
    if summary_df is None or summary_df.empty:
        return pd.DataFrame(columns=cols_out)

    blank_pass_labels = set(blank_pass_labels or [])
    bypass_labels     = set(bypass_labels     or [])
    ever_failed_by_station = ever_failed_by_station or {}

    df = summary_df.copy()
    if wo_subset:
        if isinstance(wo_subset, str):
            wo_subset = [wo_subset]
        keep = df["WO"].astype(str).str.strip().isin([w.strip() for w in wo_subset])
        df = df[keep].reset_index(drop=True)

    if df.empty:
        return pd.DataFrame(columns=cols_out)

    sn_series = df["SN"].astype(str).str.strip()
    overall_ever_failed = pd.Series(False, index=df.index)
    any_tested          = pd.Series(False, index=df.index)
    rows = []

    for label, col_name in STATION_FLOW:
        if label in bypass_labels or col_name not in df.columns:
            continue

        col_str = df[col_name].fillna("").astype(str).str.strip()
        col_up  = col_str.str.upper()

        if label in blank_pass_labels:
            tested_mask = pd.Series(True, index=df.index)
        else:
            tested_mask = col_str != ""

        # True FPY: device must not appear in this station's raw ever-failed set
        ever_failed_here  = ever_failed_by_station.get(label, set())
        raw_failed_mask   = sn_series.isin(ever_failed_here)
        first_pass_mask   = tested_mask & ~raw_failed_mask
        not_first_pass    = tested_mask & raw_failed_mask

        tested  = int(tested_mask.sum())
        fp_n    = int(first_pass_mask.sum())
        rfail_n = int(not_first_pass.sum())
        fpy_str = f"{fp_n / tested * 100:.1f}%" if tested > 0 else "—"

        overall_ever_failed |= raw_failed_mask
        any_tested          |= tested_mask

        rows.append({"Station": label, "Tested": tested,
                     "First Pass": fp_n, "Re-test Pass/Fail": rfail_n,
                     "FPY": fpy_str})

    total = len(df)
    fp    = int((~overall_ever_failed & any_tested).sum())
    rows.append({"Station": "Overall (all stations)", "Tested": total,
                 "First Pass": fp, "Re-test Pass/Fail": total - fp,
                 "FPY": f"{fp / total * 100:.1f}%" if total > 0 else "—"})

    return pd.DataFrame(rows, columns=cols_out)

def compute_wo_yield_table(summary_df: pd.DataFrame,
                           wo_subset=None,
                           blank_pass_labels=None,
                           bypass_labels=None,
                           switch_test_pass=False) -> pd.DataFrame:
    """Per-WO end-to-end yield: runs the full funnel for each WO and returns
    the Total row's Input Qty and Yield %."""
    cols_out = ["Work Order", "Input Qty", "Yield %"]
    if summary_df is None or summary_df.empty:
        return pd.DataFrame(columns=cols_out)

    df = summary_df.copy()
    if wo_subset:
        if isinstance(wo_subset, str):
            wo_subset = [wo_subset]
        keep = df["WO"].astype(str).str.strip().isin([w.strip() for w in wo_subset])
        df = df[keep].reset_index(drop=True)

    if df.empty:
        return pd.DataFrame(columns=cols_out)

    unique_wos = sorted(df["WO"].astype(str).str.strip().unique())
    kw = dict(blank_pass_labels=blank_pass_labels,
              bypass_labels=bypass_labels,
              switch_test_pass=switch_test_pass)
    rows = []
    for wo in unique_wos:
        if not wo:
            continue
        funnel = build_yield_funnel(df, wo_subset=[wo], **kw)
        total_row = funnel[funnel["Station"] == "Total"]
        if total_row.empty:
            continue
        rows.append({
            "Work Order": wo,
            "Input Qty":  int(total_row["Input Qty"].iloc[0]),
            "Yield %":    total_row["Yield"].iloc[0],
        })

    return pd.DataFrame(rows, columns=cols_out)


# ───────────────────────────────────────────────────────────────
# Qt helpers

class DataFrameModel(QAbstractTableModel):
    def __init__(self, df=pd.DataFrame(), parent=None):
        super().__init__(parent)
        self._df = df.copy()

    def setDataFrame(self, df: pd.DataFrame):
        self.beginResetModel()
        self._df = df.copy()
        self.endResetModel()

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
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return str(self._df.columns[section])
        return str(section + 1)
    
class StationListModel(DataFrameModel):
    """
    Model for the station detail list showing WO, SN and status of ALL stations.

    Any SUMMARY status column:
        - PASS      -> light green fill
        - non-blank -> light red fill
    """
    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None

        value = self._df.iat[index.row(), index.column()]

        # normal text
        if role in (Qt.DisplayRole, Qt.EditRole):
            return "" if pd.isna(value) else str(value)

        # background pass/fail colouring for all SUMMARY status columns
        if role == Qt.BackgroundRole:
            try:
                col_name = str(self._df.columns[index.column()])
            except Exception:
                col_name = ""

            if col_name in SUMMARY_STATUS_COLUMNS:
                s = "" if pd.isna(value) else str(value).strip().upper()
                if not s:
                    return None
                if s == "PASS":
                    return QBrush(QColor("#C6EFCE"))   # light green
                else:
                    return QBrush(QColor("#FFC7CE"))   # light red

        return None


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

        # normal text
        if role in (Qt.DisplayRole, Qt.EditRole):
            return "" if pd.isna(value) else str(value)

        # PASS / FAIL colouring
        if role == Qt.BackgroundRole:
            try:
                col_name = str(self._df.columns[index.column()])
            except Exception:
                col_name = ""

            if col_name in ("CH_Pass_Fail", "Pass/Fail"):
                v_str = "" if pd.isna(value) else str(value).strip().upper()
                if v_str == "PASS":
                    return QBrush(QColor("#C6EFCE"))   # light green
                elif v_str not in ("", " "):
                    return QBrush(QColor("#FFC7CE"))   # light red
            return None

        return None
class PivotDataFrameModel(DataFrameModel):
    """
    Model for pivot tables in StationPivotDialog / ThreeTempPivotDialog.

    MAX, MIN and CH1..CH8 cells are filled light red if the value
    is outside spec limits from the external spec txt.
    """
    def __init__(self, df=pd.DataFrame(), station_name: str = "", parent=None):
        super().__init__(df, parent)
        self.station_name = station_name or ""

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

            # Only colour numeric metric columns (MAX, MIN, CHx)
            if col_name in ("MAX", "MIN") or col_name.startswith("CH"):
                metric = self._df.iloc[index.row(), 0]  # first column = Metric name
                spec = get_spec_for_metric(str(metric), self.station_name)
                if spec is None:
                    return None

                lsl, usl = spec
                try:
                    v = float(value)
                except Exception:
                    return None

                out_of_spec = (
                    (lsl is not None and v < lsl) or
                    (usl is not None and v > usl)
                )
                if out_of_spec:
                    return QBrush(QColor("#FFC7CE"))  # light red fill

        return None

def _build_station_pivot(sub: pd.DataFrame, station_name: str | None = None) -> pd.DataFrame:
    """
    Build pivot table:
      Metric | MIN | MAX | CH1..CH8

    - MIN = LSL from spec file (if found)
    - MAX = USL from spec file (if found)
    - CH1..CH8 = latest measured value per channel
    - Metrics with any real data appear first, then all-NaN ones.
    """
    if sub is None or sub.empty:
        return pd.DataFrame()

    # Make sure Burn-in has CHNumber
    df = _rename_burnin_channel(sub).copy()

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
        # keep all numeric-ish metrics; order later
        _ = pd.to_numeric(df[col], errors="coerce")
        metric_cols.append(col)

    filled_rows = []
    empty_rows = []

    for metric in metric_cols:
        ch_vals = []
        for ch in range(1, 9):
            mask = df["_CH_IDX_"] == ch
            s = pd.to_numeric(df.loc[mask, metric], errors="coerce").dropna()
            v = float(s.iloc[-1]) if not s.empty else np.nan
            ch_vals.append(v)

        arr = np.array(ch_vals, dtype=float)

        # Look up spec for this metric + station
        lsl = usl = np.nan
        spec = get_spec_for_metric(str(metric), station_name)
        if spec is not None:
            lsl, usl = spec

        row = [metric, lsl, usl] + ch_vals

        if np.all(np.isnan(arr)):
            empty_rows.append(row)
        else:
            filled_rows.append(row)

    rows = filled_rows + empty_rows
    cols = ["Metric", "MIN", "MAX"] + [f"CH{i}" for i in range(1, 9)]
    return pd.DataFrame(rows, columns=cols)




def _infer_three_temp_stage(ch):
    """
    Infer RT / LT / HT from CHNumber.
    Works for patterns like: '1_RT', '1_LT', '1_HT',
    and also '1_ATS', '1_ATS_LT', '1_ATS_HT'.
    """
    if ch is None:
        return None
    s = str(ch).strip().upper()
    if not s:
        return None

    # First use the same parser as the rest of the app
    lane, stage = parse_lane_stage(s)
    if stage in ("RT", "LT", "HT"):
        return stage
    if stage == "ATS":
        # Final Test ATS (no suffix) = RT
        return "RT"

    # Fallback: look at common suffixes
    if s.endswith("_ATS") or s.endswith("_ATS_RT") or s.endswith("_RT"):
        return "RT"
    if s.endswith("_ATS_LT") or s.endswith("_LT"):
        return "LT"
    if s.endswith("_ATS_HT") or s.endswith("_HT"):
        return "HT"

    return None


def _split_three_temp_frames(raw_df: pd.DataFrame):
    """
    Given raw rows for 3TBER / TCBER / Final Test for a single SN,
    return (df_norm, stages) where stages is:
        { "RT": df_rt, "LT": df_lt, "HT": df_ht }
    and df_norm is the normalised raw dataframe.
    """
    if raw_df is None or raw_df.empty:
        empty = pd.DataFrame()
        return empty, {"RT": empty, "LT": empty, "HT": empty}

    df = normalize_test_df(raw_df).copy()

    if "CHNumber" in df.columns:
        df["_STAGE_3TEMP_"] = df["CHNumber"].apply(_infer_three_temp_stage)
    else:
        df["_STAGE_3TEMP_"] = None

    stages = {}
    for tag in ("RT", "LT", "HT"):
        stages[tag] = df[df["_STAGE_3TEMP_"] == tag].copy()

    return df, stages

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
class StationPivotDialog(QDialog):
    """
    Popup that shows:
      - Pivot view (metrics vs channels) with SID selector
      - Raw rows (with PASS/FAIL colour)
    for ONE SN + ONE station.
    """
    def __init__(self, sn: str, station_name: str, raw_df: pd.DataFrame, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"{station_name} – SN {sn}")
        self.resize(1100, 650)
        self._station_name = station_name
        self._raw_all_df = normalize_test_df(raw_df)

        lay = QVBoxLayout(self)

        info = QLabel(f"SN: {sn}    Station: {station_name}    Rows: {len(raw_df)}")
        lay.addWidget(info)

        # Test-run selector row
        sid_row = QHBoxLayout()
        sid_row.addWidget(QLabel("Test run:"))
        self._sid_combo = QComboBox()
        self._sid_combo.setMinimumWidth(220)
        self._sid_items = _build_test_run_list(self._raw_all_df)
        for label, _tn in self._sid_items:
            self._sid_combo.addItem(label)
        sid_row.addWidget(self._sid_combo)
        sid_row.addStretch(1)
        lay.addLayout(sid_row)

        tabs = QTabWidget()
        lay.addWidget(tabs, 1)

        # For Burn-in, use the same "latest per lane + cycle" reduction
        # that we use in SUMMARY, so the pivot isn't empty.
        if station_name == "Burn-in":
            pivot_source = reduce_burnin_latest_per_lane_cycle(self._raw_all_df)
        else:
            pivot_source = self._raw_all_df

        # Pivot tab (spec-coloured, MIN/MAX from spec file if available)
        pivot_df = _build_station_pivot(pivot_source, station_name=station_name)

        pivot_model = PivotDataFrameModel(pivot_df, station_name=station_name)

        self._pivot_view = QTableView()
        self._pivot_view.setModel(pivot_model)
        compact_table(self._pivot_view, row_h=26, min_col_w=90, first_col_w=160)
        self._pivot_view.setContextMenuPolicy(Qt.CustomContextMenu)
        self._pivot_view.customContextMenuRequested.connect(
            lambda pos, tv=self._pivot_view: self._save_table_as_csv(tv, f"{station_name}_pivot", pos)
        )

        tab_pivot = QWidget()
        v1 = QVBoxLayout(tab_pivot)
        v1.addWidget(self._pivot_view)
        tabs.addTab(tab_pivot, "Pivot")

        # Raw tab
        raw_model = DetailDataFrameModel(self._raw_all_df)
        raw_view = QTableView()
        raw_view.setModel(raw_model)
        compact_table(raw_view, row_h=24, min_col_w=90, first_col_w=150)
        raw_view.setContextMenuPolicy(Qt.CustomContextMenu)
        raw_view.customContextMenuRequested.connect(
            lambda pos, tv=raw_view: self._save_table_as_csv(tv, f"{station_name}_raw", pos)
        )

        tab_raw = QWidget()
        v2 = QVBoxLayout(tab_raw)
        v2.addWidget(raw_view)
        tabs.addTab(tab_raw, "Raw")

        # Connect combo after UI is built
        self._sid_combo.currentIndexChanged.connect(self._on_sid_changed)

    def _on_sid_changed(self, idx):
        if idx < 0 or idx >= len(self._sid_items):
            return
        _label, tn_val = self._sid_items[idx]
        if tn_val is None:
            if self._station_name == "Burn-in":
                filtered = reduce_burnin_latest_per_lane_cycle(self._raw_all_df)
            else:
                filtered = self._raw_all_df
        else:
            tn_col = _find_testnumber_col(self._raw_all_df)
            if tn_col:
                filtered = self._raw_all_df[
                    self._raw_all_df[tn_col].astype(str).str.strip() == tn_val
                ].copy()
            else:
                filtered = self._raw_all_df
        pivot_df = _build_station_pivot(filtered, station_name=self._station_name)
        model = PivotDataFrameModel(pivot_df, station_name=self._station_name)
        self._pivot_view.setModel(model)
        compact_table(self._pivot_view, row_h=26, min_col_w=90, first_col_w=160)

    def _save_table_as_csv(self, tv: QTableView, title: str, pos):
        menu = QMenu(tv)
        act_save = menu.addAction(f"Save '{title}' table as CSV...")
        action = menu.exec_(tv.viewport().mapToGlobal(pos))
        if action != act_save:
            return

        path, _ = QFileDialog.getSaveFileName(
            self, "Save table as CSV", f"{title}.csv",
            "CSV files (*.csv);;All files (*.*)"
        )
        if not path:
            return

        model = tv.model()
        if isinstance(model, DataFrameModel):
            df = model._df
        else:
            rows = model.rowCount()
            cols = model.columnCount()
            data = []
            headers = [model.headerData(c, Qt.Horizontal) for c in range(cols)]
            for r in range(rows):
                row_vals = [model.data(model.index(r, c), Qt.DisplayRole) for c in range(cols)]
                data.append(row_vals)
            df = pd.DataFrame(data, columns=headers)

        try:
            df.to_csv(path, index=False)
        except Exception as e:
            QMessageBox.critical(self, "Save error", str(e))

class ThreeTempPivotDialog(QDialog):
    """
    Popup for 3TBER / TCBER / Final Test (multi-temperature):
      Tabs:
        • RT – pivot (with SID selector)
        • LT – pivot
        • HT – pivot
        • Raw – all temps combined
    """
    def __init__(self, sn: str, station_name: str, raw_df: pd.DataFrame, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"{station_name} – SN {sn} (RT/LT/HT)")
        self.resize(1150, 680)
        self._station_name = station_name
        self._raw_all_df = normalize_test_df(raw_df) if raw_df is not None and not raw_df.empty else pd.DataFrame()

        layout = QVBoxLayout(self)

        info = QLabel(f"SN: {sn}    Station: {station_name}    Total rows: {len(raw_df)}")
        layout.addWidget(info)

        # Test-run selector row
        sid_row = QHBoxLayout()
        sid_row.addWidget(QLabel("Test run:"))
        self._sid_combo = QComboBox()
        self._sid_combo.setMinimumWidth(220)
        self._sid_items = _build_test_run_list(self._raw_all_df)
        for label, _tn in self._sid_items:
            self._sid_combo.addItem(label)
        sid_row.addWidget(self._sid_combo)
        sid_row.addStretch(1)
        layout.addLayout(sid_row)

        self._tabs = QTabWidget()
        layout.addWidget(self._tabs, 1)

        # Build initial pivot tabs
        self._stage_tab_widgets = {}
        self._build_pivot_tabs(self._raw_all_df)

        # Raw tab (all temps)
        raw_model = DetailDataFrameModel(self._raw_all_df)
        raw_view = QTableView()
        raw_view.setModel(raw_model)
        compact_table(raw_view, row_h=24, min_col_w=90, first_col_w=150)
        raw_view.setContextMenuPolicy(Qt.CustomContextMenu)
        raw_view.customContextMenuRequested.connect(
            lambda pos, tv=raw_view: self._save_table_as_csv(
                tv, f"{station_name}_raw", pos
            )
        )

        raw_page = QWidget()
        v2 = QVBoxLayout(raw_page)
        v2.addWidget(raw_view)
        self._tabs.addTab(raw_page, "Raw")

        # Connect SID combo after UI is built
        self._sid_combo.currentIndexChanged.connect(self._on_sid_changed)

    def _build_pivot_tabs(self, source_df):
        """Build/rebuild RT/LT/HT pivot tabs from source_df."""
        # Remove old stage tabs (keep Raw tab at end)
        for stage in list(self._stage_tab_widgets.keys()):
            for i in range(self._tabs.count()):
                if self._tabs.tabText(i) == stage:
                    self._tabs.removeTab(i)
                    break
        self._stage_tab_widgets.clear()

        _df_norm, stages = _split_three_temp_frames(source_df)

        insert_idx = 0
        any_pivot = False
        for stage in ("RT", "LT", "HT"):
            stage_df = stages.get(stage)
            if stage_df is None or stage_df.empty:
                continue

            pivot_df = _build_station_pivot(stage_df, station_name=self._station_name)
            if pivot_df.empty:
                continue

            model = PivotDataFrameModel(pivot_df, station_name=self._station_name)
            view = QTableView()
            view.setModel(model)
            compact_table(view, row_h=26, min_col_w=90, first_col_w=160)
            view.setContextMenuPolicy(Qt.CustomContextMenu)
            view.customContextMenuRequested.connect(
                lambda pos, tv=view, st=stage: self._save_table_as_csv(
                    tv, f"{self._station_name}_{st}_pivot", pos
                )
            )

            page = QWidget()
            v = QVBoxLayout(page)
            v.addWidget(QLabel(f"{stage}: {len(stage_df)} raw rows \u2192 {len(pivot_df)} metrics"))
            v.addWidget(view)

            self._tabs.insertTab(insert_idx, page, stage)
            self._stage_tab_widgets[stage] = page
            insert_idx += 1
            any_pivot = True

        if not any_pivot:
            page = QWidget()
            v = QVBoxLayout(page)
            v.addWidget(QLabel("No RT/LT/HT rows found for this device at this station."))
            self._tabs.insertTab(0, page, "Pivot")
            self._stage_tab_widgets["Pivot"] = page

    def _on_sid_changed(self, idx):
        if idx < 0 or idx >= len(self._sid_items):
            return
        _label, tn_val = self._sid_items[idx]
        if tn_val is None:
            source = self._raw_all_df
        else:
            tn_col = _find_testnumber_col(self._raw_all_df)
            if tn_col:
                source = self._raw_all_df[
                    self._raw_all_df[tn_col].astype(str).str.strip() == tn_val
                ].copy()
            else:
                source = self._raw_all_df
        self._build_pivot_tabs(source)

    def _save_table_as_csv(self, tv: QTableView, title: str, pos):
        menu = QMenu(tv)
        act_save = menu.addAction(f"Save '{title}' table as CSV...")
        action = menu.exec_(tv.viewport().mapToGlobal(pos))
        if action != act_save:
            return

        path, _ = QFileDialog.getSaveFileName(
            self, "Save table as CSV", f"{title}.csv",
            "CSV files (*.csv);;All files (*.*)"
        )
        if not path:
            return

        model = tv.model()
        if isinstance(model, DataFrameModel):
            df = model._df
        else:
            rows = model.rowCount()
            cols = model.columnCount()
            data = []
            headers = [model.headerData(c, Qt.Horizontal) for c in range(cols)]
            for r in range(rows):
                row_vals = [model.data(model.index(r, c), Qt.DisplayRole) for c in range(cols)]
                data.append(row_vals)
            df = pd.DataFrame(data, columns=headers)

        try:
            df.to_csv(path, index=False)
        except Exception as e:
            QMessageBox.critical(self, "Save error", str(e))

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

def compact_table(tv: QTableView, row_h=26, min_col_w=100, first_col_w=140):
    tv.setAlternatingRowColors(True)
    tv.setWordWrap(False)
    tv.setTextElideMode(Qt.ElideRight)
    tv.setItemDelegate(ElideDelegate(row_h=row_h, parent=tv))
    vh = tv.verticalHeader()
    vh.setVisible(True)
    vh.setDefaultSectionSize(row_h)
    hh = tv.horizontalHeader()
    hh.setStretchLastSection(False)
    hh.setMinimumSectionSize(60)
    hh.setDefaultSectionSize(min_col_w)
    hh.setSectionResizeMode(QHeaderView.Interactive)
    tv.setMinimumHeight(305)
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
    

def build_app_qss(font_px=13):
    return f"""
* {{
  font-family: 'Segoe UI','Inter','Arial';
  font-size: {font_px}px;
}}
QWidget {{
  background: #FFFFFF;
  color: #111827;
}}
QHeaderView::section {{
  background: #F2F4F7;
  color: #111827;
  padding: 4px 8px;
  border: 1px solid #E2E6EC;
  font-weight: 600;
}}
QTableView {{
  background: #FFFFFF;
  gridline-color: #E2E6EC;
  alternate-background-color: #F9FAFB;
}}
QTableView::item:selected {{
  background: #E7F0FF;
  color: #111827;
}}
QPushButton {{
  background: #FFFFFF;
  border: 1px solid #D9DEE7;
  border-radius: 6px;
  padding: 6px 12px;
}}
QPushButton:hover {{
  background: #F8FAFF;
  border-color: #C7D3EA;
}}
QTextEdit {{
  background: #FFFFFF;
  border: 1px solid #D9DEE7;
  border-radius: 6px;
}}
QProgressBar {{
  background: #EFF3FA;
  border: 1px solid #E0E5EF;
  border-radius: 6px;
  text-align: center;
}}
QProgressBar::chunk {{
  background-color: #5AA8FF;
  border-radius: 6px;
}}
"""

# ───────────────────────────────────────────────────────────────
# Workers

def _collect_ever_failed(raw_df: pd.DataFrame) -> set:
    """Return COMPONENTIDs that had ANY fail row in the raw (all-attempts) data.

    Handles two column conventions:
      - CH_Pass_Fail / Pass/Fail  (TRX, BURNIN, BER, Switch tables)
      - Status / Result / PassFail (Mode Hopping table)
    """
    if raw_df is None or raw_df.empty or "COMPONENTID" not in raw_df.columns:
        return set()

    comp = raw_df["COMPONENTID"].astype(str).str.strip()

    # TRX / BURNIN / BER / Switch: CH_Pass_Fail or Pass/Fail column
    if "CH_Pass_Fail" in raw_df.columns or "Pass/Fail" in raw_df.columns:
        df = normalize_test_df(raw_df)
        fail_mask = df["CH_Pass_Fail"].astype(str).str.upper().str.strip() != "PASS"
        return set(comp[fail_mask].unique())

    # Mode Hopping and similar: Status / Result column
    st_col = _find_ci_col(raw_df, ["Status", "Result", "PassFail", "Pass_Fail", "PASSFAIL"])
    if st_col:
        fail_mask = raw_df[st_col].astype(str).str.upper().str.strip() != "PASS"
        return set(comp[fail_mask].unique())

    return set()


class SummaryWorker(QThread):
    progressPct = pyqtSignal(int)
    done = pyqtSignal(pd.DataFrame)
    error = pyqtSignal(str)

    def __init__(self, source_mode, wo_list, device_codes,
                 use_date_filter, start_ymd, end_ymd,
                 sn_list=None, sn_batch_name="SN Batch", sn_wo_map=None, parent=None):
        super().__init__(parent)
        self.source_mode = source_mode  # "WO" | "DEVICE" | "SN"
        self.wo_list = [w.strip() for w in wo_list if w.strip()]
        self.device_codes = [d.strip() for d in device_codes if d.strip()]
        self.use_date_filter = use_date_filter
        self.start_ymd = start_ymd
        self.end_ymd = end_ymd

        self.sn_list = [str(s).strip() for s in (sn_list or []) if str(s).strip()]
        self.sn_batch_name = (sn_batch_name or "SN Batch").strip() or "SN Batch"
        self.sn_wo_map = sn_wo_map
    def run(self):
        try:
            pct = 0
            def bump(add):
                nonlocal pct
                pct = min(100, pct + add)
                self.progressPct.emit(pct)

            bump(3)  # connect
            conn = pyodbc.connect(DB_CONN)

            # components + WO map
            if self.source_mode == "WO":
                if not self.wo_list:
                    raise RuntimeError("Please define at least one Work Order.")
                comps = fetch_components_by_wos(conn, self.wo_list)
                if not comps:
                    raise RuntimeError("No COMPONENTIDs found for these WO(s).")
                wo_map = fetch_wo_map(conn, comps, restrict_wos=self.wo_list)
            elif self.source_mode == "SN":
                if not self.sn_list:
                    raise RuntimeError("Please paste at least one SN (COMPONENTID).")
                comps = self.sn_list
                if self.sn_wo_map:
                    wo_map = self.sn_wo_map
                else:
                    wo_map = fetch_wo_map(conn, comps)
                    for c in comps:
                        key = str(c).strip()
                        if not wo_map.get(key):
                            wo_map[key] = self.sn_batch_name
            else:
                if not self.device_codes:
                    raise RuntimeError("Please select at least one Device (DR8+/FR4).")
                comps, wo_map = fetch_components_by_devices(conn, self.device_codes)
                if not comps:
                    raise RuntimeError("No COMPONENTIDs found for selected Device ID(s).")
            bump(7)

            # TRX (main) — fetch raw first, capture ever-failed per type, then dedup
            trx_df_raw = fetch_table_for_components_bulk(conn, TRX_TABLE, comps)
            _trx_raw_split = split_trx_types_raw(trx_df_raw)
            ever_failed_by_station = {
                "DDMI Cal":   _collect_ever_failed(_trx_raw_split.get("DDMI",  pd.DataFrame())),
                "TP2TP3 RT":  _collect_ever_failed(_trx_raw_split.get("RT",    pd.DataFrame())),
                "TP2TP3 LT":  _collect_ever_failed(_trx_raw_split.get("LT",    pd.DataFrame())),
                "TP2TP3 HT":  _collect_ever_failed(_trx_raw_split.get("HT",    pd.DataFrame())),
                "Final Test": _collect_ever_failed(_trx_raw_split.get("FINAL", pd.DataFrame())),
            }
            trx_df = keep_latest_per_channel(trx_df_raw) if not trx_df_raw.empty else trx_df_raw
            trx_split = split_trx_types(trx_df)
            type_to_frames: dict = {"TRX": trx_split}
            bump(10)

            # FW Writing + Mode Hopping — raw for FPY, then dedup for summary
            fw_df_raw = fetch_table_for_components_bulk(conn, FWWRITE_TABLE, comps)
            ever_failed_by_station["FW Writing"] = _collect_ever_failed(fw_df_raw)
            fw_df = reduce_fw_latest_per_component(fw_df_raw) if not fw_df_raw.empty else fw_df_raw
            type_to_frames["FWWRITE"] = fw_df

            mh_df_raw = fetch_table_for_components_bulk(conn, MODEHOP_TABLE, comps)
            ever_failed_by_station["Mode Hopping"] = _collect_ever_failed(mh_df_raw)
            if not mh_df_raw.empty:
                mh_latest_map = fetch_latest_modehop_testnumber_map(conn, comps)
                mh_df = reduce_modehop_to_master_latest(mh_df_raw, mh_latest_map)
                type_to_frames["MODEHOP"] = mh_df
            bump(6)

            # Other tables — raw for FPY, then dedup for summary
            _key_to_station = {"BURNIN": "Burn-in", "SWITCH": "Switch Test",
                               "3TBER": "3T Test",  "TCBER":  "TCBER"}
            for key, table in TABLES.items():
                if key == "BURNIN":
                    df_raw = fetch_burnin_combined(conn, comps)
                elif key == "SWITCH":
                    df_raw = fetch_switch_combined(conn, comps)
                else:
                    df_raw = fetch_table_for_components_bulk(conn, table, comps)
                ever_failed_by_station[_key_to_station[key]] = _collect_ever_failed(df_raw)
                df_all = df_raw.copy() if not df_raw.empty else df_raw
                if not df_all.empty:
                    if key in {"TCBER", "3TBER"}:
                        df_all = reduce_ber_latest_per_lane_stage(df_all)
                    elif key == "BURNIN":
                        df_all = reduce_burnin_latest_per_lane_cycle(df_all)
                    else:
                        df_all = keep_latest_per_channel(df_all)
                type_to_frames[key] = df_all
                bump(8)

            self.ever_failed_by_station = ever_failed_by_station

            # Optional DDMI date filter
            if self.use_date_filter:
                ddmi_date_map = compute_ddmi_date_for_components(trx_split, comps)
                filtered = []
                for c in comps:
                    key = str(c).strip()
                    d = ddmi_date_map.get(key)
                    if d is not None and self.start_ymd <= d <= self.end_ymd:
                        filtered.append(c)
                if not filtered:
                    raise RuntimeError("No devices found within this DDMI date range.")
                comps = filtered

            bump(15)
            summary_df = build_summary(type_to_frames, comps, wo_map)
            bump(32)

            try:
                conn.close()
            except Exception:
                pass

            self.progressPct.emit(100)
            self.done.emit(summary_df)

        except Exception as e:
            tb = traceback.format_exc(limit=2)
            self.error.emit(f"{e}\n{tb}")

class ScheduleWorker(QThread):
    progressPct = pyqtSignal(int)
    done = pyqtSignal(pd.DataFrame)
    error = pyqtSignal(str)

    def __init__(self, target_ymd: int, device_ids: list[str], parent=None):
        super().__init__(parent)
        self.target_ymd = target_ymd
        self.device_ids = [d.strip() for d in device_ids if d.strip()]

    def run(self):
        try:
            pct = 0
            def bump(add):
                nonlocal pct
                pct = min(100, pct + add)
                self.progressPct.emit(pct)

            conn = pyodbc.connect(DB_CONN)
            bump(5)

            # Fetch each station's records directly by SID date — no master-table lookup needed.
            trx_df       = fetch_table_for_date(conn, TRX_TABLE,          self.target_ymd)
            trx_split_raw = split_trx_types_raw(trx_df)
            bump(20)

            fw_df        = fetch_table_for_date(conn, FWWRITE_TABLE,       self.target_ymd)
            mh_df        = fetch_table_for_date(conn, MODEHOP_TABLE,       self.target_ymd)
            bump(10)

            burnin_df    = fetch_burnin_for_date(conn, self.target_ymd)
            ber3t_df     = fetch_table_for_date(conn, TABLES["3TBER"],     self.target_ymd)
            tcber_df     = fetch_table_for_date(conn, TABLES["TCBER"],     self.target_ymd)
            switch_df    = fetch_switch_for_date(conn, self.target_ymd)
            bump(30)

            def count_df(df: pd.DataFrame) -> int:
                if df is None or df.empty or "COMPONENTID" not in df.columns:
                    return 0
                return df["COMPONENTID"].astype(str).str.strip().nunique()

            rows = []
            station_dfs = [
                ("DDMI Cal",     trx_split_raw["DDMI"]),
                ("FW Writing",   fw_df),
                ("TP2TP3 RT",    trx_split_raw["RT"]),
                ("TP2TP3 LT",    trx_split_raw["LT"]),
                ("TP2TP3 HT",    trx_split_raw["HT"]),
                ("Mode Hopping", mh_df),
                ("Burn-in",      burnin_df),
                ("3T Test",      ber3t_df),
                ("TCBER",        tcber_df),
                ("Final Test",   trx_split_raw["FINAL"]),
                ("Switch Test",  switch_df),
            ]

            for label, df in station_dfs:
                qty = count_df(df)
                if qty > 0:
                    rows.append({
                        "Station": label,
                        "Cycle Time": STATION_CYCLE.get(label, ""),
                        "Slots": STATION_SLOTS.get(label, ""),
                        "Tested Qty": int(qty),
                    })

            bump(35)

            schedule_df = pd.DataFrame(rows, columns=["Station","Cycle Time","Slots","Tested Qty"])

            # Order schedule by real process flow
            order_map = {label: idx for idx, (label, _) in enumerate(STATION_FLOW)}
            schedule_df["_order"] = schedule_df["Station"].map(order_map).fillna(9999).astype(int)
            schedule_df = (
                schedule_df
                .sort_values("_order")
                .drop(columns="_order")
                .reset_index(drop=True)
            )

            try:
                conn.close()
            except Exception:
                pass

            self.progressPct.emit(100)
            self.done.emit(schedule_df)

        except Exception as e:
            tb = traceback.format_exc(limit=2)
            self.error.emit(f"{e}\n{tb}")


# ───────────────────────────────────────────────────────────────

# Weekly Yield (WO discovery) helpers

def _ymd_int_to_date(ymd: int):
    """Convert YYYYMMDD int to datetime.date (returns None if invalid)."""
    try:
        y = int(ymd) // 10000
        m = (int(ymd) // 100) % 100
        d = int(ymd) % 100
        return date(y, m, d)
    except Exception:
        return None


class WeeklyWOIndexWorker(QThread):
    """Discover WOs received in a date range from MASTER.LIV_MASTER_SID and group into ISO weeks."""

    done = pyqtSignal(list, str)   # (wo_lists, title)
    error = pyqtSignal(str)

    def __init__(self, start_ymd: int, end_ymd: int, device_type: str = "", parent=None):
        super().__init__(parent)
        self.start_ymd = int(start_ymd)
        self.end_ymd = int(end_ymd)
        self.device_type = (device_type or "").strip()

    def run(self):
        try:
            conn = pyodbc.connect(DB_CONN)
            cur = conn.cursor()

            # MASTER-based WO receipt date:
            # LIV_MASTER_SID examples: A2026011311071825858 -> date = SUBSTRING(SID,2,8)
            sql = f"""
            WITH wo_first AS (
                SELECT
                    LTRIM(RTRIM(m.WO)) AS WO,
                    MIN(
                        TRY_CONVERT(int,
                            CASE
                                WHEN LEFT(LTRIM(RTRIM(CONVERT(varchar(80), m.LIV_MASTER_SID))), 1) = 'A'
                                    THEN SUBSTRING(LTRIM(RTRIM(CONVERT(varchar(80), m.LIV_MASTER_SID))), 2, 8)
                                ELSE LEFT(LTRIM(RTRIM(CONVERT(varchar(80), m.LIV_MASTER_SID))), 8)
                            END
                        )
                    ) AS FirstYMD
                FROM {MASTER_WO_TABLE} m
                WHERE m.WO IS NOT NULL
                  AND m.LIV_MASTER_SID IS NOT NULL
                  AND (? = '' OR m.DEVICETYPE = ?)
                GROUP BY LTRIM(RTRIM(m.WO))
            )
            SELECT WO, FirstYMD
            FROM wo_first
            WHERE FirstYMD BETWEEN ? AND ?
            ORDER BY FirstYMD, WO
            """

            cur.execute(sql, (self.device_type, self.device_type, self.start_ymd, self.end_ymd))
            rows = cur.fetchall()

            wo_first = []
            for r in rows:
                if not r:
                    continue
                wo = str(r[0]).strip() if r[0] is not None else ""
                if not wo:
                    continue
                try:
                    first_ymd = int(r[1]) if r[1] is not None else None
                except Exception:
                    first_ymd = None
                if not first_ymd:
                    continue
                wo_first.append((wo, first_ymd))

            try:
                conn.close()
            except Exception:
                pass

            # Group WOs by ISO week using FirstReceivedDate
            buckets = {}  # (iso_year, iso_week) -> {"wos": set(), "week_start": date, "week_end": date}
            for wo, first_ymd in wo_first:
                dt = _ymd_int_to_date(first_ymd)
                if not dt:
                    continue
                iso_year, iso_week, iso_wday = dt.isocalendar()
                week_start = dt - timedelta(days=iso_wday - 1)
                week_end = week_start + timedelta(days=6)

                key = (iso_year, iso_week)
                if key not in buckets:
                    buckets[key] = {"wos": set(), "week_start": week_start, "week_end": week_end}
                buckets[key]["wos"].add(wo)

            wo_lists = []
            for (iso_year, iso_week) in sorted(buckets.keys()):
                info = buckets[(iso_year, iso_week)]
                wos = sorted(info["wos"])
                ws = info["week_start"]
                we = info["week_end"]
                name = f"{iso_year%100:02d}-w{iso_week:02d}"
                wo_lists.append({"name": name, "wos": wos})

            title = f"Weekly Yield – WO Received ({self.start_ymd}–{self.end_ymd})"
            self.done.emit(wo_lists, title)

        except Exception as e:
            tb = traceback.format_exc(limit=4)
            self.error.emit(f"{e}\n{tb}")


class WeeklyYieldRangeDialog(QDialog):
    """Pick Year + Date Range + optional DeviceType. Produces week tabs for WOs received (MASTER.LIV_MASTER_SID)."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Weekly Yield (WO received from MASTER)")
        self.setMinimumWidth(420)

        # Year dropdown
        today = QDate.currentDate()
        y = today.year()
        self.cbo_year = QComboBox()
        for yy in [y - 2, y - 1, y, y + 1]:
            self.cbo_year.addItem(str(yy), yy)
        self.cbo_year.setCurrentText(str(y))

        self.date_from = QDateEdit()
        self.date_to = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_to.setCalendarPopup(True)

        # Optional DeviceType filter
        self.txt_devtype = QLineEdit()
        self.txt_devtype.setPlaceholderText("Optional (e.g., TTX)")

        # default full-year (selected year)
        self._apply_year_defaults()
        self.cbo_year.currentIndexChanged.connect(self._apply_year_defaults)

        self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)

        form = QFormLayout()
        form.addRow("Year:", self.cbo_year)
        form.addRow("From date:", self.date_from)
        form.addRow("To date:", self.date_to)
        form.addRow("DeviceType:", self.txt_devtype)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(self.buttons)

    def _apply_year_defaults(self):
        yy = int(self.cbo_year.currentData())
        self.date_from.setDate(QDate(yy, 1, 1))
        self.date_to.setDate(QDate(yy, 12, 31))

    def get_values(self):
        d1 = self.date_from.date()
        d2 = self.date_to.date()
        if d2 < d1:
            d1, d2 = d2, d1
        start_ymd = int(d1.toString("yyyyMMdd"))
        end_ymd = int(d2.toString("yyyyMMdd"))
        devtype = self.txt_devtype.text().strip()
        return devtype, start_ymd, end_ymd



# ───────────────────────────────────────────────────────────────
# Excel Report Export (all Yield tabs)


def _safe_sheet_name(name: str) -> str:
    """Excel sheet name max 31 chars; remove invalid chars."""
    name = (name or "").strip()
    name = re.sub(r"[\\/\?\*\[\]:]", "-", name)
    if not name:
        name = "Report"
    return name[:31]


def _master_first_date_map(conn, wos: list[str], device_type: str = "") -> dict[str, int]:
    """Return {WO: FirstYMD_int} from MASTER.LIV_MASTER_SID for provided WOs."""
    wos = [str(w).strip() for w in (wos or []) if str(w).strip()]
    if not wos:
        return {}
    cur = conn.cursor()
    out: dict[str, int] = {}
    chunk = 800  # keep params manageable
    for i in range(0, len(wos), chunk):
        sub = wos[i:i+chunk]
        placeholders = ",".join(["?"] * len(sub))
        sql = f"""
        SELECT
            LTRIM(RTRIM(m.WO)) AS WO,
            MIN(
                TRY_CONVERT(int,
                    CASE
                        WHEN LEFT(LTRIM(RTRIM(CONVERT(varchar(80), m.LIV_MASTER_SID))), 1) = 'A'
                            THEN SUBSTRING(LTRIM(RTRIM(CONVERT(varchar(80), m.LIV_MASTER_SID))), 2, 8)
                        ELSE LEFT(LTRIM(RTRIM(CONVERT(varchar(80), m.LIV_MASTER_SID))), 8)
                    END
                )
            ) AS FirstYMD
        FROM {MASTER_WO_TABLE} m
        WHERE m.WO IN ({placeholders})
          AND m.LIV_MASTER_SID IS NOT NULL
          AND (? = '' OR m.DEVICETYPE = ?)
        GROUP BY LTRIM(RTRIM(m.WO))
        """
        params = list(sub) + [device_type, device_type]
        cur.execute(sql, params)
        for wo, ymd in cur.fetchall():
            if wo is None or ymd is None:
                continue
            try:
                out[str(wo).strip()] = int(ymd)
            except Exception:
                pass
    return out


class ReportExportWorker(QThread):
    """Export all current Yield tabs to a formatted Excel report."""
    progress = pyqtSignal(int, str)  # pct, message
    done = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self,
                 out_path: str,
                 tab_lists: list[dict],
                 summary_df: pd.DataFrame,
                 fail_modes: list[str],
                 blank_pass_labels: set,
                 bypass_labels: set,
                 device_type: str = "",
                 parent=None):
        super().__init__(parent)
        self.out_path = out_path
        self.tab_lists = tab_lists or []
        self.summary_df = summary_df.copy() if summary_df is not None else pd.DataFrame()
        self.fail_modes = list(fail_modes or [])
        self.blank_pass_labels = set(blank_pass_labels or set())
        self.bypass_labels = set(bypass_labels or set())
        self.device_type = (device_type or "").strip()

    def run(self):
        try:
            if self.summary_df.empty or not self.tab_lists:
                raise RuntimeError("No data to export. Run Get Data first.")

            # Workbook init
            wb = Workbook()
            try:
                wb.remove(wb.active)
            except Exception:
                pass

            # Styles
            hdr_fill = PatternFill("solid", fgColor="1F4E79")
            hdr_font = Font(color="FFFFFF", bold=True)
            thin = Side(style="thin", color="A0A0A0")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)
            center = Alignment(horizontal="center", vertical="center", wrap_text=True)

            self.progress.emit(1, "Fetching WO start dates...")
            conn = pyodbc.connect(DB_CONN)
            all_wos = sorted({wo for lst in self.tab_lists for wo in (lst.get("wos") or [])})
            wo_first_map = _master_first_date_map(conn, all_wos, self.device_type)
            try:
                conn.close()
            except Exception:
                pass

            total_wos = sum(len(lst.get("wos") or []) for lst in self.tab_lists)
            done_wos = 0

            for t_idx, lst in enumerate(self.tab_lists, start=1):
                tab_name = str(lst.get("name") or f"Tab{t_idx}")
                wos = [str(w).strip() for w in (lst.get("wos") or []) if str(w).strip()]
                if not wos:
                    continue

                self.progress.emit(int(5 + (t_idx-1)/max(1,len(self.tab_lists))*10), f"Building sheet {tab_name}...")
                ws = wb.create_sheet(_safe_sheet_name(tab_name))

                # Title
                ws["A1"] = tab_name
                ws["A1"].font = Font(bold=True, size=14)

                # Left WO summary table
                headers = ["WO", "StartDate", "WOTY", "FINISHQTY", "FAILQTY", "WIPQTY", "Week", "Yield"]
                for c, h in enumerate(headers, start=1):
                    cell = ws.cell(row=3, column=c, value=h)
                    cell.fill = hdr_fill
                    cell.font = hdr_font
                    cell.alignment = center
                    cell.border = border

                r = 4
                for wo in wos:
                    funnel = build_yield_funnel(self.summary_df, wo_subset=[wo],
                                                blank_pass_labels=self.blank_pass_labels,
                                                bypass_labels=self.bypass_labels)
                    tot = funnel[funnel["Station"] == "Total"]
                    if tot.empty:
                        woty = finish = fail = wip = 0
                        yld = ""
                    else:
                        rr = tot.iloc[0]
                        woty = int(rr.get("Input Qty", 0) or 0)
                        finish = int(rr.get("Output Qty", 0) or 0)
                        fail = int(rr.get("Fail Qty", 0) or 0)
                        wip = int(rr.get("WIP", 0) or 0)
                        yld = str(rr.get("Yield", "") or "")

                    ymd = wo_first_map.get(wo)
                    sd = _ymd_int_to_date(int(ymd)) if ymd else None
                    sd_str = sd.strftime("%m/%d/%Y") if sd else ""

                    row_vals = [wo, sd_str, woty, finish, fail, wip, tab_name, yld]
                    for c, v in enumerate(row_vals, start=1):
                        cell = ws.cell(row=r, column=c, value=v)
                        cell.border = border
                        if c >= 3:
                            cell.alignment = center
                    r += 1

                    done_wos += 1
                    if total_wos:
                        self.progress.emit(min(95, int(10 + done_wos/total_wos*60)),
                                           f"Exporting WOs ({done_wos}/{total_wos})...")

                # Totals row
                ws.cell(row=r, column=1, value="Totals").font = Font(bold=True)
                for c in range(1, 9):
                    ws.cell(row=r, column=c).border = border

                for c in range(3, 7):
                    col = get_column_letter(c)
                    cell = ws.cell(row=r, column=c, value=f"=SUM({col}4:{col}{r-1})")
                    cell.font = Font(bold=True)
                    cell.alignment = center
                    cell.border = border

                ycell = ws.cell(row=r, column=8, value=f"=IF((D{r}+E{r})=0,"",D{r}/(D{r}+E{r}))")
                ycell.number_format = "0.00%"
                ycell.font = Font(bold=True)
                ycell.alignment = center
                ycell.border = border

                # column widths
                widths = [14, 12, 8, 10, 9, 8, 10, 10]
                for i, w in enumerate(widths, start=1):
                    ws.column_dimensions[get_column_letter(i)].width = w

                # Pareto table + chart (combined)
                pareto_col = 10  # J
                ws.cell(row=1, column=pareto_col, value="Fail Pareto (Combined)").font = Font(bold=True)

                df_tab = self.summary_df[self.summary_df["WO"].astype(str).str.strip().isin(wos)].reset_index(drop=True)
                records = compute_first_fail_records(df_tab,
                                                     fail_modes=self.fail_modes,
                                                     blank_pass_labels=self.blank_pass_labels,
                                                     bypass_labels=self.bypass_labels)
                labels = []
                for st, fm in records:
                    if not st or not fm:
                        continue
                    station_tag = re.sub(r"[\s\-]+", "", st)
                    labels.append(f"{fm}@{station_tag}")

                if labels:
                    ser = pd.Series(labels)
                    counts = ser.value_counts().sort_values(ascending=False)
                    cats = counts.index.tolist()
                    vals = counts.values.astype(int).tolist()
                    total_fail = sum(vals) if vals else 0

                    p_headers = ["FailMode@Station", "Qty", "%"]
                    for j, h in enumerate(p_headers):
                        cell = ws.cell(row=3, column=pareto_col + j, value=h)
                        cell.fill = hdr_fill
                        cell.font = hdr_font
                        cell.alignment = center
                        cell.border = border

                    pr = 4
                    max_items = min(20, len(cats))
                    for k in range(max_items):
                        ws.cell(row=pr, column=pareto_col, value=cats[k]).border = border
                        qcell = ws.cell(row=pr, column=pareto_col + 1, value=int(vals[k])); qcell.border = border; qcell.alignment = center
                        pcell = ws.cell(row=pr, column=pareto_col + 2, value=(vals[k]/total_fail if total_fail else "")); pcell.border = border
                        pcell.number_format = "0.0%"
                        pcell.alignment = center
                        pr += 1

                    # Chart image via matplotlib (bar + cumulative line)
                    try:
                        import matplotlib
                        matplotlib.use("Agg")
                        import matplotlib.pyplot as plt
                        fig = plt.figure(figsize=(7.4, 3.0))
                        ax = fig.add_subplot(111)
                        x = list(range(max_items))
                        ax.bar(x, vals[:max_items])
                        ax.set_xticks(x)
                        ax.set_xticklabels(cats[:max_items], rotation=45, ha="right", fontsize=7)
                        ax.set_ylabel("Qty")
                        ax2 = ax.twinx()
                        cum = np.cumsum(vals[:max_items])
                        cum_pct = (cum / total_fail * 100.0) if total_fail else np.zeros_like(cum)
                        ax2.plot(x, cum_pct, marker='o')
                        ax2.set_ylim(0, 110)
                        ax2.set_ylabel("Cum %")
                        fig.tight_layout()
                        tmp_png = os.path.join(tempfile.gettempdir(), f"pareto_{_safe_sheet_name(tab_name)}.png")
                        fig.savefig(tmp_png, dpi=160, bbox_inches="tight")
                        plt.close(fig)
                        img = XLImage(tmp_png)
                        img.anchor = f"J{pr+1}"
                        ws.add_image(img)
                    except Exception:
                        pass

                # Right-side per-WO station tables
                base_col = 12  # L
                base_row = 3
                per_row = 3
                block_w = 6
                block_h = len(STATION_FLOW) + 4

                for idx2, wo in enumerate(wos):
                    r0 = base_row + (idx2 // per_row) * (block_h + 2)
                    c0 = base_col + (idx2 % per_row) * (block_w + 2)

                    ws.cell(row=r0, column=c0, value=f"{wo}").font = Font(bold=True)

                    hdrs2 = ["Station", "In", "Out", "Fail", "WIP", "Yield"]
                    for j, h in enumerate(hdrs2):
                        cell = ws.cell(row=r0+1, column=c0+j, value=h)
                        cell.fill = hdr_fill
                        cell.font = hdr_font
                        cell.alignment = center
                        cell.border = border

                    funnel = build_yield_funnel(self.summary_df, wo_subset=[wo],
                                                blank_pass_labels=self.blank_pass_labels,
                                                bypass_labels=self.bypass_labels)
                    rr2 = r0 + 2
                    for _, fr in funnel.iterrows():
                        st = fr.get("Station", "")
                        if not st:
                            continue
                        vals2 = [
                            st,
                            int(fr.get("Input Qty", 0) or 0),
                            int(fr.get("Output Qty", 0) or 0),
                            int(fr.get("Fail Qty", 0) or 0),
                            int(fr.get("WIP", 0) or 0),
                            fr.get("Yield", ""),
                        ]
                        for j, v in enumerate(vals2):
                            cell = ws.cell(row=rr2, column=c0+j, value=v)
                            cell.border = border
                            cell.alignment = center if j > 0 else Alignment(horizontal="left", vertical="center")
                        rr2 += 1

                ws.freeze_panes = "A4"

            self.progress.emit(98, "Saving Excel...")
            wb.save(self.out_path)
            self.progress.emit(100, "Done")
            self.done.emit(self.out_path)

        except Exception as e:
            tb = traceback.format_exc(limit=6)
            self.error.emit(f"{e}\n{tb}")


class ScheduleDetailWorker(QThread):
    progress = pyqtSignal(int, str)
    done = pyqtSignal(str, str, pd.DataFrame, pd.DataFrame)  # station_label, detail_kind, detail_df, summary_df
    error = pyqtSignal(str)

    def __init__(self, target_ymd: int, station_label: str, summary_df: pd.DataFrame, parent=None):
        super().__init__(parent)
        self.target_ymd = target_ymd
        self.station_label = station_label
        self.summary_df = summary_df.copy() if summary_df is not None else pd.DataFrame()

    def run(self):
        try:
            conn = pyodbc.connect(DB_CONN)
            
            # 1) Which devices (COMPONENTIDs) were tested at this station on this date?
            comps, wo_map_device = fetch_components_by_devices(conn, SCHEDULE_DEVICE_IDS)
            if not comps:
                raise RuntimeError("No COMPONENTIDs found for Schedule devices.")

            # Decide which table to use for the "tested" set
            station_df = pd.DataFrame()

            if self.station_label in ("DDMI Cal", "FW Writing", "TP2TP3 RT", "TP2TP3 LT", "TP2TP3 HT", "Mode Hopping", "Final Test"):
                trx_df = fetch_table_for_components_bulk(conn, TRX_TABLE, comps)
                trx_split_raw = split_trx_types_raw(trx_df)

                if self.station_label == "DDMI Cal": station_df = trx_split_raw["DDMI"]
                elif self.station_label == "FW Writing": station_df = trx_split_raw.get("FW", pd.DataFrame())
                elif self.station_label == "TP2TP3 RT": station_df = trx_split_raw["RT"]
                elif self.station_label == "TP2TP3 LT": station_df = trx_split_raw["LT"]
                elif self.station_label == "TP2TP3 HT": station_df = trx_split_raw["HT"]
                elif self.station_label == "Mode Hopping": station_df = trx_split_raw.get("MODEHOP", pd.DataFrame())
                elif self.station_label == "Final Test": station_df = trx_split_raw["FINAL"]
            elif self.station_label == "Burn-in": station_df = fetch_burnin_combined(conn, comps)
            elif self.station_label == "3T Test": station_df = fetch_table_for_components_bulk(conn, TABLES["3TBER"], comps)
            elif self.station_label == "TCBER": station_df = fetch_table_for_components_bulk(conn, TABLES["TCBER"], comps)
            elif self.station_label == "Switch Test": station_df = fetch_switch_combined(conn, comps)

            if station_df is None or station_df.empty or "SID" not in station_df.columns:
                raise RuntimeError(f"No test rows for {self.station_label} on this date.")

            station_df = station_df.copy()
            station_df["_SID_DATE_"] = station_df["SID"].apply(sid_to_yyyymmdd_int)
            station_day = station_df[station_df["_SID_DATE_"] == self.target_ymd]
            if station_day.empty:
                raise RuntimeError(f"No {self.station_label} rows on selected date.")

            sns = station_day["COMPONENTID"].astype(str).str.strip().dropna()
            unique_sns = sorted(set(sns))
            if not unique_sns:
                raise RuntimeError(f"No COMPONENTIDs for {self.station_label} on this date.")

            # 2) Build full SUMMARY for these devices
            wo_map = {}
            if self.summary_df is not None and not self.summary_df.empty and "SN" in self.summary_df.columns and "WO" in self.summary_df.columns:
                _tmp = self.summary_df.copy()
                _tmp["SN"] = _tmp["SN"].astype(str).str.strip()
                _tmp["WO"] = _tmp["WO"].astype(str).str.strip()
                wo_map.update(dict(zip(_tmp["SN"], _tmp["WO"])))

            for sn in unique_sns:
                if sn not in wo_map or not wo_map.get(sn, "").strip():
                    wo_map[sn] = str(wo_map_device.get(sn, "") or "").strip()

            type_to_frames = {}
            trx_df_all = fetch_table_for_components_bulk(conn, TRX_TABLE, unique_sns)
            if not trx_df_all.empty: trx_df_all = keep_latest_per_channel(trx_df_all)
            type_to_frames["TRX"] = split_trx_types(trx_df_all)

            for key, table in TABLES.items():
                if key == "BURNIN":
                    df_all = fetch_burnin_combined(conn, unique_sns)
                elif key == "SWITCH":
                    df_all = fetch_switch_combined(conn, unique_sns)
                else:
                    df_all = fetch_table_for_components_bulk(conn, table, unique_sns)
                if not df_all.empty:
                    if key in {"TCBER", "3TBER"}: df_all = reduce_ber_latest_per_lane_stage(df_all)
                    elif key == "BURNIN": df_all = reduce_burnin_latest_per_lane_cycle(df_all)
                    else: df_all = keep_latest_per_channel(df_all)
                type_to_frames[key] = df_all
            
            fw_df = fetch_table_for_components_bulk(conn, FWWRITE_TABLE, unique_sns)
            if not fw_df.empty: fw_df = reduce_fw_latest_per_component(fw_df)
            type_to_frames["FWWRITE"] = fw_df
            
            mh_df = fetch_table_for_components_bulk(conn, MODEHOP_TABLE, unique_sns)
            if not mh_df.empty:
                mh_latest_map = fetch_latest_modehop_testnumber_map(conn, unique_sns)
                mh_df = reduce_modehop_to_master_latest(mh_df, mh_latest_map)
            type_to_frames["MODEHOP"] = mh_df

            summary_df = build_summary(type_to_frames, unique_sns, wo_map)
            if summary_df.empty:
                raise RuntimeError("SUMMARY is empty for these Schedule devices.")

            summary_df["SN"] = summary_df["SN"].astype(str).str.strip()
            summary_df = summary_df[summary_df["SN"].isin(unique_sns)].reset_index(drop=True)

            station_cols = [c for (c, _sub) in SUMMARY_COLUMNS if c in summary_df.columns]
            cols = ["WO", "SN"] + station_cols
            detail_df = summary_df.loc[:, cols].copy()
            
            detail_kind = f"Tested Qty (date {self.target_ymd})"
            self.done.emit(self.station_label, detail_kind, detail_df, summary_df)

        except Exception as e:
            self.error.emit(str(e))
        finally:
            try:
                conn.close()
            except Exception:
                pass


# ───────────────────────────────────────────────────────────────
# Dialogs (WO lists & Fail modes)

class ManageListsDialog(QDialog):
    def __init__(self, wo_lists=None, sn_lists=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Manage WO & SN Lists")
        self.setMinimumSize(650, 420)

        self.main_tabs = QTabWidget()

        # WO Page
        wo_page = QWidget()
        wo_layout = QVBoxLayout(wo_page)
        self.wo_tabs = QTabWidget()

        btn_add_wo = QPushButton("Add WO List")
        btn_add_wo.clicked.connect(lambda: self.add_wo_list_tab())
        btn_import_auto_wo = QPushButton("Import Auto WOs from DB")
        btn_import_auto_wo.setToolTip(
            "Fetch WOs using the same Device filter as Yield → Auto WO by Device,\n"
            "then create one list tab per WO (List name = WO)."
        )
        btn_import_auto_wo.clicked.connect(self.import_auto_wos_from_db)
        
        wo_btn_row = QHBoxLayout()
        wo_btn_row.addWidget(btn_add_wo)
        wo_btn_row.addWidget(btn_import_auto_wo)
        wo_btn_row.addStretch(1)
        wo_layout.addLayout(wo_btn_row)
        wo_layout.addWidget(self.wo_tabs)
        
        # SN Page
        sn_page = QWidget()
        sn_layout = QVBoxLayout(sn_page)
        self.sn_tabs = QTabWidget()

        btn_add_sn = QPushButton("Add SN List")
        btn_add_sn.clicked.connect(lambda: self.add_sn_list_tab())
        
        sn_btn_row = QHBoxLayout()
        sn_btn_row.addWidget(btn_add_sn)
        sn_btn_row.addStretch(1)
        sn_layout.addLayout(sn_btn_row)
        sn_layout.addWidget(self.sn_tabs)

        self.main_tabs.addTab(wo_page, "WO Lists")
        self.main_tabs.addTab(sn_page, "SN Lists")

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addWidget(self.main_tabs)
        layout.addWidget(buttons)

        if wo_lists:
            for lst in wo_lists:
                name = lst.get("name", "")
                wos  = lst.get("wos", [])
                text = "\n".join([str(x) for x in wos if str(x).strip()])
                self.add_wo_list_tab(name, text)
        else:
            self.add_wo_list_tab("List 1", "")

        if sn_lists:
            for lst in sn_lists:
                name = lst.get("name", "")
                sns  = lst.get("sns", [])
                text = "\n".join([str(x) for x in sns if str(x).strip()])
                self.add_sn_list_tab(name, text)
        else:
            self.add_sn_list_tab("SN List 1", "")

    def add_wo_list_tab(self, name="", text=""):
        page = QWidget()
        v = QVBoxLayout(page)

        row_name = QHBoxLayout()
        row_name.addWidget(QLabel("List name:"))
        name_edit = QLineEdit(name)
        row_name.addWidget(name_edit)
        v.addLayout(row_name)

        v.addWidget(QLabel("Work Orders (one per line, or comma separated):"))
        txt = QTextEdit()
        txt.setPlainText(text)
        v.addWidget(txt)

        btn_del = QPushButton("Delete this list")
        btn_del.clicked.connect(lambda: self.remove_wo_tab(page))
        v.addWidget(btn_del)

        v.addStretch(1)

        page.name_edit = name_edit
        page.wo_edit = txt

        idx = self.wo_tabs.addTab(page, name or f"List {self.wo_tabs.count() + 1}")
        self.wo_tabs.setCurrentIndex(idx)

    def add_sn_list_tab(self, name="", text=""):
        page = QWidget()
        v = QVBoxLayout(page)

        row_name = QHBoxLayout()
        row_name.addWidget(QLabel("List name:"))
        name_edit = QLineEdit(name)
        row_name.addWidget(name_edit)
        v.addLayout(row_name)

        v.addWidget(QLabel("Serial Numbers (one per line, or comma separated):"))
        txt = QTextEdit()
        txt.setPlainText(text)
        v.addWidget(txt)

        btn_del = QPushButton("Delete this list")
        btn_del.clicked.connect(lambda: self.remove_sn_tab(page))
        v.addWidget(btn_del)

        v.addStretch(1)

        page.name_edit = name_edit
        page.sn_edit = txt

        idx = self.sn_tabs.addTab(page, name or f"List {self.sn_tabs.count() + 1}")
        self.sn_tabs.setCurrentIndex(idx)

    def remove_wo_tab(self, page):
        idx = self.wo_tabs.indexOf(page)
        if idx >= 0:
            self.wo_tabs.removeTab(idx)
        if self.wo_tabs.count() == 0:
            self.add_wo_list_tab("List 1", "")

    def remove_sn_tab(self, page):
        idx = self.sn_tabs.indexOf(page)
        if idx >= 0:
            self.sn_tabs.removeTab(idx)
        if self.sn_tabs.count() == 0:
            self.add_sn_list_tab("SN List 1", "")

    def get_wo_lists(self):
        out = []
        for i in range(self.wo_tabs.count()):
            page = self.wo_tabs.widget(i)
            name = page.name_edit.text().strip() or f"List {i+1}"
            raw  = page.wo_edit.toPlainText().strip()
            wos = [p.strip() for line in raw.splitlines()
                   for p in line.split(",") if p.strip()]
            if not wos:
                continue
            out.append({"name": name, "wos": wos})
        return out

    def get_sn_lists(self):
        out = []
        for i in range(self.sn_tabs.count()):
            page = self.sn_tabs.widget(i)
            name = page.name_edit.text().strip() or f"SN List {i+1}"
            raw  = page.sn_edit.toPlainText().strip()
            sns = [p.strip() for line in raw.splitlines()
                   for p in line.split(",") if p.strip()]
            if not sns:
                continue
            out.append({"name": name, "sns": sns})
        return out

    def accept(self):
        if not self.get_wo_lists() and not self.get_sn_lists():
            QMessageBox.warning(self, "No Lists", "Please enter at least one WO or SN in a list.")
            return
        super().accept()

    def _get_auto_device_ids(self) -> list[str]:
        """Use the same Device IDs as Yield → Auto WO by Device (DR8+/FR4).

        If opened from Yield window, it uses chk_dr8/chk_fr4 state.
        If none selected (or dialog used standalone), defaults to BOTH.
        """
        ids: list[str] = []
        p = self.parent()
        try:
            if p is not None and hasattr(p, "chk_dr8") and hasattr(p, "chk_fr4"):
                if p.chk_dr8.isChecked():
                    ids.append(DR8_DEVICE_ID)
                if p.chk_fr4.isChecked():
                    ids.append(FR4_DEVICE_ID)
        except Exception:
            ids = []

        if not ids:
            ids = [DR8_DEVICE_ID, FR4_DEVICE_ID]
        return ids

    def import_auto_wos_from_db(self):
        device_ids = self._get_auto_device_ids()
        if not device_ids:
            QMessageBox.warning(self, "No devices", "No Device IDs selected for Auto WO import.")
            return

        try:
            conn = pyodbc.connect(DB_CONN)
        except Exception as e:
            QMessageBox.critical(self, "DB connect error", str(e))
            return

        try:
            wos = fetch_wos_by_devices(conn, device_ids)
        except Exception as e:
            QMessageBox.critical(self, "DB query error", str(e))
            return
        finally:
            try:
                conn.close()
            except Exception:
                pass

        if not wos:
            QMessageBox.information(self, "No WOs", "No WOs found for selected Auto devices.")
            return

        if len(wos) > 250:
            ret = QMessageBox.question(
                self,
                "Large import",
                f"This will create {len(wos)} list tabs (one per WO).\n\nContinue?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if ret != QMessageBox.Yes:
                return

        # Existing names (avoid duplicates)
        existing = set()
        for i in range(self.wo_tabs.count()):
            page = self.wo_tabs.widget(i)
            try:
                nm = page.name_edit.text().strip()
                if nm:
                    existing.add(nm)
            except Exception:
                pass
            try:
                t = self.wo_tabs.tabText(i).strip()
                if t:
                    existing.add(t)
            except Exception:
                pass

        added = 0
        skipped = 0
        for wo in wos:
            name = str(wo).strip()
            if not name:
                continue
            if name in existing:
                skipped += 1
                continue
            self.add_wo_list_tab(name, name)
            existing.add(name)
            added += 1

        QMessageBox.information(
            self,
            "Import complete",
            f"Imported WOs using Auto device filter.\n\nAdded: {added}\nSkipped (already existed): {skipped}"
        )


class FailModeSNDialog(QDialog):
    """Popup table listing all devices (WO/SN) for a chosen Pareto bar."""
    def __init__(self, title: str, rows: list[dict], parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumSize(900, 520)

        self.table = QTableWidget(self)
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["WO", "SN", "Station", "FailMode", "FailText"])
        self.table.setRowCount(len(rows))

        for r, d in enumerate(rows):
            self.table.setItem(r, 0, QTableWidgetItem(str(d.get("WO",""))))
            self.table.setItem(r, 1, QTableWidgetItem(str(d.get("SN",""))))
            self.table.setItem(r, 2, QTableWidgetItem(str(d.get("Station",""))))
            self.table.setItem(r, 3, QTableWidgetItem(str(d.get("FailMode",""))))
            self.table.setItem(r, 4, QTableWidgetItem(str(d.get("FailText",""))))

        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.resizeColumnsToContents()

        btns = QDialogButtonBox(QDialogButtonBox.Close, parent=self)
        btns.rejected.connect(self.reject)
        btns.accepted.connect(self.accept)

        lay = QVBoxLayout(self)
        lay.addWidget(self.table, 1)
        lay.addWidget(btns)




class FailModesDialog(QDialog):
    def __init__(self, modes=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Manage Fail Modes (keywords)")
        self.setMinimumSize(500, 350)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(
            "Enter fail modes (keywords), one per line.\n"
            "Example: TCOrder, Sen, ModeHopping, RLM, VCC_offset"
        ))

        self.txt = QTextEdit()
        if modes:
            self.txt.setPlainText("\n".join(modes))
        layout.addWidget(self.txt)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_modes(self):
        raw = self.txt.toPlainText().strip()
        modes = [line.strip() for line in raw.splitlines() if line.strip()]
        return modes

    def accept(self):
        if not self.get_modes():
            QMessageBox.warning(self, "No fail modes", "Please enter at least one fail-mode keyword.")
            return
        super().accept()

# ───────────────────────────────────────────────────────────────
# Main window
def compute_first_fail_records(summary_df: pd.DataFrame,
                               fail_modes: list[str],
                               blank_pass_labels=None,
                               bypass_labels=None) -> list[tuple[str, str]]:
    """
    For each device (row) compute its first failing station (by STATION_FLOW)
    and classify the fail text into a fail-mode keyword.

    Returns list of (first_fail_station_label, fail_mode_string).
    """
    if summary_df is None or summary_df.empty:
        return []

    if blank_pass_labels is None:
        blank_pass_labels = set()
    else:
        blank_pass_labels = set(blank_pass_labels)

    # Normalise bypass station list
    if bypass_labels is None:
        bypass_labels = set()
    else:
        bypass_labels = set(bypass_labels)

    df = summary_df.copy()

    fm_list = [m.strip() for m in fail_modes if m.strip()]
    fm_upper = [m.upper() for m in fm_list]

    records: list[tuple[str, str]] = []

    for _, row in df.iterrows():
        first_fail_label = None
        fail_text = None

        # Walk stations in process order
        for label, col_name in STATION_FLOW:
            if label in bypass_labels:
                continue

            if col_name not in df.columns:
                continue

            val = row[col_name]
            s = "" if pd.isna(val) else str(val).strip()

            if s == "":
                if label in blank_pass_labels:
                    # blank treated as PASS
                    continue
                else:
                    # no data yet, chain stops
                    break

            s_up = s.upper()
            if s_up == "PASS":
                continue

            first_fail_label = label
            fail_text = s
            break

        if first_fail_label is None or not fail_text:
            continue

        t_up = fail_text.upper()
        fail_mode = "OTHER"
        for mode, mode_up in zip(fm_list, fm_upper):
            if mode_up in t_up:
                fail_mode = mode
                break

        records.append((first_fail_label, fail_mode))

    return records


def compute_first_fail_details(summary_df: pd.DataFrame,
                               fail_modes: list[str],
                               blank_pass_labels=None,
                               bypass_labels=None) -> list[dict]:
    """
    Like compute_first_fail_records, but returns per-device details:
      {WO,SN,Station,FailMode,FailText}
    """
    if summary_df is None or summary_df.empty:
        return []

    if blank_pass_labels is None:
        blank_pass_labels = set()
    else:
        blank_pass_labels = {str(x).strip().upper() for x in blank_pass_labels if str(x).strip()}

    if bypass_labels is None:
        bypass_labels = set()
    else:
        bypass_labels = {str(x).strip().upper() for x in bypass_labels if str(x).strip()}

    modes_u = [str(m).strip().upper() for m in (fail_modes or []) if str(m).strip()]
    details = []

    for _, row in summary_df.iterrows():
        wo = str(row.get("WO", "")).strip()
        sn = str(row.get("SN", "")).strip()

        first_fail_label = None
        first_fail_text  = ""

        for label, col in STATION_FLOW:
            val = str(row.get(col, "")).strip()
            if not val:
                continue
            vup = val.upper()
            if vup in blank_pass_labels:
                continue
            if vup in bypass_labels:
                continue
            if vup == "PASS":
                continue
            first_fail_label = label
            first_fail_text  = val
            break

        if first_fail_label is None:
            continue

        # classify fail mode
        t_up = (first_fail_text or "").upper()
        fail_mode = "OTHER"
        for mode, mode_up in zip((fail_modes or []), modes_u):
            if mode_up and mode_up in t_up:
                fail_mode = str(mode).strip()
                break

        details.append({
            "WO": wo,
            "SN": sn,
            "Station": first_fail_label,
            "FailMode": fail_mode,
            "FailText": first_fail_text
        })

    return details

class StationListDialog(QDialog):
    """
    Shows:
      - table with WO, SN and status of ALL stations (SUMMARY columns)
      - Pareto chart for the chosen station's fails

    Devices are the ones in detail_df; Pareto is computed on that subset.
    """
    def __init__(self, host_window, station_label: str, detail_kind: str,
                 detail_df: pd.DataFrame, summary_context_df: pd.DataFrame,
                 fail_modes: list[str], blank_pass_labels, bypass_labels=None, parent=None):
        super().__init__(parent)
        self._host = host_window
        self.station_label = station_label
        self.detail_kind = detail_kind

        # IMPORTANT: reset index so Qt row index (0..N-1) matches DataFrame
        self.detail_df = detail_df.reset_index(drop=True).copy()
        self.summary_context_df = summary_context_df.copy()
        self.fail_modes = list(fail_modes)
        self.blank_pass_labels = set(blank_pass_labels) if blank_pass_labels else set()
        self.bypass_labels = set(bypass_labels) if bypass_labels else set()

        self.setWindowTitle(f"{station_label} – {detail_kind}")
        self.resize(1000, 650)

        lay = QVBoxLayout(self)

        info = QLabel(
            f"Station: {station_label}    View: {detail_kind}    Devices: {len(self.detail_df)}"
        )
        lay.addWidget(info)

        # Table: WO, SN + ALL station status columns
        self.model = StationListModel(self.detail_df)
        self.view = QTableView()
        self.view.setModel(self.model)
        compact_table(self.view, row_h=26, min_col_w=110, first_col_w=140)
        

        self.view.doubleClicked.connect(self.on_table_double_clicked)
        self.view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.view.customContextMenuRequested.connect(self._on_table_context_menu)

        lay.addWidget(self.view, 2)

        # Pareto chart
        self.fig = Figure(figsize=(7, 3.2))
        self.canvas = FigureCanvas(self.fig)
        self.canvas.setMinimumHeight(320)
        self.canvas.setContextMenuPolicy(Qt.CustomContextMenu)
        self.canvas.customContextMenuRequested.connect(self._on_canvas_context_menu)
        lay.addWidget(self.canvas, 1)

        self.build_pareto()

    def on_table_double_clicked(self, index):
        """
        Double-click any row:
          - on WO/SN columns      -> use this dialog's station_label
          - on any station column -> use that station (mapped from column name)
        """
        if not index.isValid():
            return

        col_name = str(self.detail_df.columns[index.column()])
        row = self.detail_df.iloc[index.row()]

        sn = str(row.get("SN", "")).strip()
        if not sn:
            return

        # Which station this cell refers to?
        if col_name in ("WO", "SN"):
            station_label = self.station_label
        else:
            station_label = COL_TO_LABEL.get(col_name)
            if not station_label:
                # clicked a non-station column (e.g. WO) other than SN/WO -> ignore
                return

        self._host.open_station_pivot_for_sn(sn, station_label)

    # ---------- Context menus for table & chart ----------

    def _on_table_context_menu(self, pos):
        menu = QMenu(self.view)
        act_csv = menu.addAction("Save this table as CSV...")
        act_png = menu.addAction("Save this table as PNG image...")
        act = menu.exec_(self.view.viewport().mapToGlobal(pos))

        if act == act_csv:
            path, _ = QFileDialog.getSaveFileName(
                self,
                "Save table as CSV",
                "station_status_table.csv",
                "CSV files (*.csv);;All files (*.*)",
            )
            if path:
                try:
                    self.model._df.to_csv(path, index=False)
                except Exception as e:
                    QMessageBox.critical(self, "Save error", str(e))

        elif act == act_png:
            path, _ = QFileDialog.getSaveFileName(
                self,
                "Save table as PNG",
                "station_status_table.png",
                "PNG files (*.png);;All files (*.*)",
            )
            if path:
                try:
                    pix = self.view.grab()
                    pix.save(path)
                except Exception as e:
                    QMessageBox.critical(self, "Save error", str(e))

    def _on_canvas_context_menu(self, pos):
        menu = QMenu(self.canvas)
        act_png = menu.addAction("Save Pareto chart as PNG...")
        act = menu.exec_(self.canvas.mapToGlobal(pos))
        if act == act_png:
            path, _ = QFileDialog.getSaveFileName(
                self,
                "Save chart as PNG",
                "station_pareto.png",
                "PNG files (*.png);;All files (*.*)",
            )
            if path:
                try:
                    self.fig.savefig(path, dpi=150, bbox_inches="tight")
                except Exception as e:
                    QMessageBox.critical(self, "Save error", str(e))

    # ---------- Pareto ----------

    def build_pareto(self):
        fig = self.fig
        fig.clear()

        if self.summary_context_df is None or self.summary_context_df.empty or self.detail_df.empty:
            self.canvas.draw()
            return

        # Restrict summary to devices visible in this dialog
        sns_set = set(self.detail_df["SN"].astype(str).str.strip())
        df = self.summary_context_df.copy()
        df["SN"] = df["SN"].astype(str).str.strip()
        df = df[df["SN"].isin(sns_set)]

        if df.empty:
            self.canvas.draw()
            return

        # First-fail records on this subset
        details = compute_first_fail_details(df, self.fail_modes, self.blank_pass_labels, self.bypass_labels)
        station = self.station_label

        labels = [d.get("FailMode") for d in details if d.get("Station") == station and d.get("FailMode")]
        cat_to_rows = {}
        for d in details:
            if d.get("Station") == station:
                key = d.get("FailMode")
                if key:
                    cat_to_rows.setdefault(key, []).append(d)
        if not labels:
            ax = fig.add_subplot(111)
            ax.text(
                0.5,
                0.5,
                "No fails for this station\nin current device list.",
                ha="center",
                va="center",
                fontsize=11,
            )
            ax.set_axis_off()
            fig.tight_layout()
            self.canvas.draw()
            return

        ser = pd.Series(labels)
        counts = ser.value_counts().sort_values(ascending=False)

        vals = counts.values.astype(int)
        cats = counts.index.tolist()
        cum_pct = vals.cumsum() / vals.sum() * 100.0

        ax = fig.add_subplot(111)
        x = np.arange(len(cats))

        bars = bars = ax.bar(x, vals)
        for b in bars:
            try:
                b.set_picker(5)
            except Exception:
                pass

        # Make bars interactive
        for b in bars:
            try:
                b.set_picker(5)
            except Exception:
                pass
        ax.set_xlabel("Fail Mode")
        ax.set_ylabel("Qty")
        ax.set_xticks(x)
        ax.set_xticklabels(cats, rotation=45, ha="right")

        max_val = max(vals) if len(vals) > 0 else 0
        offset = max_val * 0.05 if max_val > 0 else 0.5
        for xi, v in zip(x, vals):
            ax.text(xi, v + offset, str(v), ha="center", va="bottom", fontsize=9)

        ax2 = ax.twinx()
        ax2.plot(x, cum_pct, marker="o")
        ax2.set_ylim(0, 110)
        ax2.set_ylabel("Cumulative %")
        for xi, p in zip(x, cum_pct):
            ax2.text(xi, p + 2, f"{p:.0f}%", ha="center", va="bottom", fontsize=8)

        ax.setTitle = ax.set_title(f"{station} – Fail Pareto (devices in this list)")
        # Double-click a bar to open SN list popup
        try:
            if hasattr(self, "_pareto_pick_cid") and self._pareto_pick_cid:
                self.canvas.mpl_disconnect(self._pareto_pick_cid)
        except Exception:
            pass

        def _on_pick(event):
            try:
                mouse_event = getattr(event, "mouseevent", None)
                if mouse_event is None or not mouse_event.dblclick:
                    return

                artist = event.artist
                if artist not in bars:
                    return

                # Find the index of the picked bar
                for i, bar in enumerate(bars):
                    if bar == artist:
                        cat = cats[i]
                        rows = cat_to_rows.get(cat, [])
                        title = f"{station} – {cat} – {len(rows)} devices"
                        dlg = FailModeSNDialog(title, rows, self)
                        dlg.exec_()
                        break
            except Exception:
                return
        
        self._pareto_pick_cid = self.canvas.mpl_connect("pick_event", _on_pick)

        fig.tight_layout()
        self.canvas.draw()


class YieldFunnelWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Analysis – Yield Funnel + Pareto + Schedule (trail0)")
        self.setMinimumSize(1200, 800)
        self.setStyleSheet(build_app_qss())

        self._summary_df = pd.DataFrame()
        self._ever_failed_by_station: dict = {}

        # Load persisted Yield config (WO lists + fail modes) if available
        self._wo_lists = load_saved_wo_lists(SAVED_LISTS_XLSX)
        self._sn_lists = load_saved_sn_lists(SAVED_LISTS_XLSX)
        saved_fm = load_saved_fail_modes(SAVED_LISTS_XLSX)

        # start with hard-coded default fail modes (override if saved exists)
        self._fail_modes = saved_fm if saved_fm else list(DEFAULT_FAIL_MODES)

        # multi-station blank-as-PASS
        self._blank_pass_labels = set()

        # bypass station(s): treat selected station as Input=Output, no fails, no gating
        self._bypass_labels = set()
        self._switch_test_pass = False

        self._source_mode = "WO"
        self._schedule_df = pd.DataFrame()

        # ── central scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self.setCentralWidget(scroll)

        container = QWidget()
        scroll.setWidget(container)
        layout = QVBoxLayout(container)
        layout.setSpacing(6)

        # Top row
        self.cmb_source = QComboBox()
        self.cmb_source.addItems(["WO Lists", "Auto WOs (by device)", "SN List (paste)"])

        self.btn_manage_lists = QPushButton("Manage Lists")
        self.btn_existing_lists = QPushButton("Existing Lists")
        self.btn_save_lists = QPushButton("Save Lists")
        self.btn_fail_modes = QPushButton("Fail Modes")
        self.btn_get = QPushButton("Get Data")
        self.btn_weekly = QPushButton("Weekly Yield")
        self.btn_blank_override = QPushButton("Blanks as PASS...")
        self.btn_bypass = QPushButton("Bypass...")
        self.btn_report = QPushButton("Save Report...")

        self.chk_dr8 = QCheckBox("DR8+")
        self.chk_fr4 = QCheckBox("FR4")
        self.chk_dr8.setVisible(False)
        self.chk_fr4.setVisible(False)

        # Date filter controls (moved into top row, swapped with buttons)
        self.chk_date_filter = QCheckBox("Enable date filter (DDMI date)")
        self.date_from = QDateEdit()
        self.date_to = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_to.setCalendarPopup(True)
        today = QDate.currentDate()
        self.date_from.setDate(today.addDays(-7))
        self.date_to.setDate(today)

        # Top row: Source + WO lists + Fail modes + Type + Date filter (right side)
        top_row = QHBoxLayout()
        top_row.addWidget(QLabel("Source:"))
        top_row.addWidget(self.cmb_source)
        top_row.addSpacing(10)
        top_row.addWidget(self.btn_manage_lists)
        top_row.addSpacing(6)
        top_row.addWidget(self.btn_existing_lists)
        top_row.addSpacing(6)
        top_row.addWidget(self.btn_save_lists)
        top_row.addSpacing(20)
        #top_row.addWidget(QLabel("Fail Modes:"))
        top_row.addWidget(self.btn_fail_modes)
        top_row.addSpacing(20)
        top_row.addWidget(QLabel("Type"))
        top_row.addWidget(self.chk_dr8)
        top_row.addWidget(self.chk_fr4)
        top_row.addStretch(1)
        top_row.addWidget(self.chk_date_filter)
        top_row.addSpacing(10)
        top_row.addWidget(QLabel("From:"))
        top_row.addWidget(self.date_from)
        top_row.addSpacing(10)
        top_row.addWidget(QLabel("To:"))
        top_row.addWidget(self.date_to)
        layout.addLayout(top_row)

        # Info row (unchanged)
        if self._wo_lists:
            names = ", ".join(lst.get("name", "") for lst in self._wo_lists if lst.get("name"))
            self.lbl_lists = QLabel(f"Configured lists: {names}" if names else "Configured lists: (loaded)")
        else:
            self.lbl_lists = QLabel("Configured lists: None")
        n_defaults = len(self._fail_modes)
        self.lbl_fail_modes = QLabel(f"{n_defaults} fail modes configured")
        self.lbl_override = QLabel("Blank-pass stations: None")
        self.lbl_bypass = QLabel("Bypass station: None")

        info_row = QHBoxLayout()
        info_row.addWidget(self.lbl_lists)
        info_row.addSpacing(20)
        info_row.addWidget(self.lbl_fail_modes)
        info_row.addSpacing(20)
        info_row.addWidget(self.lbl_override)
        info_row.addSpacing(20)
        info_row.addWidget(self.lbl_bypass)
        info_row.addStretch(1)
        layout.addLayout(info_row)

        # Buttons row (Get Data + Blanks as PASS...), now under info row
        btn_row = QHBoxLayout()
        btn_row.addWidget(self.btn_weekly)
        btn_row.addWidget(self.btn_get)
        btn_row.addSpacing(10)
        btn_row.addWidget(self.btn_blank_override)
        btn_row.addSpacing(10)
        btn_row.addWidget(self.btn_bypass)
        btn_row.addSpacing(10)
        btn_row.addWidget(self.btn_report)
        btn_row.addStretch(1)
        layout.addLayout(btn_row)


        # Progress bar
        self.prog = QProgressBar()
        self.prog.setRange(0, 100)
        self.prog.setVisible(False)
        layout.addWidget(self.prog)

        # Tabs (Yield + Pareto)
        self.tabs = QTabWidget()
        self.tabs.setMinimumHeight(500)
        layout.addWidget(self.tabs)

        # Schedule section
        schedule_group = QGroupBox("Schedule ")
        sg_layout = QVBoxLayout(schedule_group)

        s_row = QHBoxLayout()
        self.schedule_date = QDateEdit()
        self.schedule_date.setCalendarPopup(True)
        self.schedule_date.setDate(today)
        self.btn_schedule = QPushButton("Get Schedule")
        s_row.addWidget(QLabel("Date:"))
        s_row.addWidget(self.schedule_date)
        s_row.addSpacing(10)
        s_row.addWidget(self.btn_schedule)
        s_row.addStretch(1)
        sg_layout.addLayout(s_row)

        self.schedule_model = DataFrameModel(pd.DataFrame(columns=["Station","Cycle Time","Slots","Tested Qty"]))
        self.schedule_table = QTableView()
        self.schedule_table.setModel(self.schedule_model)
        compact_table(self.schedule_table, row_h=24, min_col_w=100, first_col_w=140)
        self.schedule_table.setMinimumHeight(420)
        self.schedule_table.setMaximumHeight(900)
        self.schedule_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.schedule_table.customContextMenuRequested.connect(
            lambda pos, t=self.schedule_table, title="Schedule": self.save_table_context_menu(pos, t, title)
        )

        self.schedule_table.doubleClicked.connect(self.on_schedule_table_double_clicked)

        sg_layout.addWidget(self.schedule_table)

        layout.addWidget(schedule_group)
        #layout.addStretch(1)

        # Signals
        self.cmb_source.activated.connect(self.on_source_changed)
        self.btn_manage_lists.clicked.connect(self.on_manage_lists)
        self.btn_existing_lists.clicked.connect(self.on_existing_lists)
        self.btn_save_lists.clicked.connect(self.on_save_lists)
        self.btn_fail_modes.clicked.connect(self.on_manage_fail_modes)
        self.btn_get.clicked.connect(self.on_get_clicked)
        self.btn_weekly.clicked.connect(self.on_weekly_clicked)
        self.btn_blank_override.clicked.connect(self.on_blank_override)
        self.btn_bypass.clicked.connect(self.on_bypass_clicked)
        self.btn_report.clicked.connect(self.on_save_report_clicked)
        self.btn_schedule.clicked.connect(self.on_get_schedule_clicked)

        self.on_source_changed(0)
    def on_yield_table_double_clicked(self, index, page_scroll, tab_title):
        if not index.isValid():
            return

        model = page_scroll.table_model
        if not isinstance(model, DataFrameModel):
            return

        df_funnel = model._df
        if df_funnel.empty:
            return

        row_idx = index.row()
        col_idx = index.column()
        if row_idx < 0 or col_idx < 0:
            return

        col_name = str(df_funnel.columns[col_idx])
        row = df_funnel.iloc[row_idx]
        station_label = str(row.get("Station", ""))

        if station_label == "Total" or not station_label:
            return

        # Which quantity column was double-clicked?
        if col_name == "Output Qty":
            detail_kind = "Output Qty (PASS devices)"
            target_kind = "pass"
        elif col_name == "Fail Qty":
            detail_kind = "Fail Qty (FAIL devices)"
            target_kind = "fail"
        elif col_name == "WIP":
            detail_kind = "WIP (no result yet)"
            target_kind = "wip"
        else:
            # Only these three columns are interactive
            return

        if self._summary_df is None or self._summary_df.empty:
            QMessageBox.warning(self, "No SUMMARY", "Run 'Get Data' first.")
            return

        # Summary for this tab (WO-list or Auto)
        df = self._summary_df.copy()
        if getattr(page_scroll, "wo_list", None):
            wo_list = [w.strip() for w in page_scroll.wo_list if w.strip()]
            df = df[df["WO"].astype(str).str.strip().isin(wo_list)].reset_index(drop=True)

        if df.empty:
            QMessageBox.information(self, "No data", "No SUMMARY rows for this tab.")
            return

        station_masks = compute_station_masks(df, self._blank_pass_labels, self._bypass_labels)
        info = station_masks.get(station_label)
        if info is None:
            QMessageBox.information(self, "No data", f"No summary for station '{station_label}'.")
            return

        if target_kind == "pass":
            mask = info["pass"]
        elif target_kind == "fail":
            mask = info["fail"]
        else:
            mask = info["wip"]

        if not mask.any():
            QMessageBox.information(
                self,
                "No devices",
                f"No devices in this tab for '{station_label}' with this quantity.",
            )
            return

        # Build detail dataframe: WO, SN + ALL SUMMARY status columns
        station_cols = [c for (c, _sub) in SUMMARY_COLUMNS if c in df.columns]
        cols = ["WO", "SN"] + station_cols
        detail_df = df.loc[mask, cols].copy()

        dlg = StationListDialog(
            host_window=self,
            station_label=station_label,
            detail_kind=detail_kind,
            detail_df=detail_df,
            summary_context_df=df,
            fail_modes=self._fail_modes,
            blank_pass_labels=self._blank_pass_labels,
            bypass_labels=self._bypass_labels,
            parent=self,
        )
        dlg.exec_()

    def on_fpy_table_double_clicked(self, index, fpy_tv, wo_list):
        if not index.isValid():
            return
        fpy_df = fpy_tv.model()._df
        if fpy_df is None or fpy_df.empty:
            return

        col_name     = str(fpy_df.columns[index.column()])
        station_label = str(fpy_df.iloc[index.row()].get("Station", ""))

        if col_name not in ("First Pass", "Re-test Pass/Fail"):
            return
        if self._summary_df is None or self._summary_df.empty:
            QMessageBox.warning(self, "No data", "Run 'Get Data' first.")
            return

        df = self._summary_df.copy()
        if wo_list:
            df = df[df["WO"].astype(str).str.strip().isin(
                [w.strip() for w in wo_list])].reset_index(drop=True)
        if df.empty:
            QMessageBox.information(self, "No data", "No rows for this tab.")
            return

        sn_col = df["SN"].astype(str).str.strip()

        if station_label == "Overall (all stations)":
            all_ever_failed = set().union(*self._ever_failed_by_station.values()) \
                if self._ever_failed_by_station else set()
            stat_cols = [col for (_, col) in STATION_FLOW if col in df.columns]
            tested_mask = df[stat_cols].fillna("").astype(str).apply(
                lambda c: c.str.strip() != "").any(axis=1)
            if col_name == "First Pass":
                mask       = tested_mask & ~sn_col.isin(all_ever_failed)
                detail_kind = "Overall — First Pass (never failed at any station)"
            else:
                mask       = tested_mask & sn_col.isin(all_ever_failed)
                detail_kind = "Overall — Re-test (failed at least once somewhere)"
        else:
            summary_col = LABEL_TO_COL.get(station_label, station_label)
            if summary_col not in df.columns:
                QMessageBox.information(self, "No data", f"No column for '{station_label}'.")
                return
            tested_mask      = df[summary_col].fillna("").astype(str).str.strip() != ""
            ever_failed_here = self._ever_failed_by_station.get(station_label, set())
            if col_name == "First Pass":
                mask       = tested_mask & ~sn_col.isin(ever_failed_here)
                detail_kind = f"{station_label} — First Pass (no prior fails)"
            else:
                mask       = tested_mask & sn_col.isin(ever_failed_here)
                detail_kind = f"{station_label} — Re-test (had at least one fail)"

        if not mask.any():
            QMessageBox.information(self, "No devices", "No devices match this selection.")
            return

        station_cols = [c for (c, _) in SUMMARY_COLUMNS if c in df.columns]
        detail_df    = df.loc[mask, ["WO", "SN"] + station_cols].copy()

        dlg = StationListDialog(
            host_window=self,
            station_label=station_label,
            detail_kind=detail_kind,
            detail_df=detail_df,
            summary_context_df=df,
            fail_modes=self._fail_modes,
            blank_pass_labels=self._blank_pass_labels,
            bypass_labels=self._bypass_labels,
            parent=self,
        )
        dlg.exec_()

    def on_schedule_table_double_clicked(self, index):
        if not index.isValid():
            return

        df_sched = self.schedule_model._df
        if df_sched is None or df_sched.empty:
            return

        row_idx = index.row()
        col_idx = index.column()
        if row_idx < 0 or col_idx < 0:
            return

        col_name = str(df_sched.columns[col_idx])
        if col_name != "Tested Qty":
            return

        row = df_sched.iloc[row_idx]
        station_label = str(row.get("Station", ""))
        qty = row.get("Tested Qty", 0)

        if not station_label or not isinstance(qty, (int, np.integer)) or qty <= 0:
            return

        target_ymd = int(self.schedule_date.date().toString("yyyyMMdd"))

        self._start_busy()
        self.schedule_detail_worker = ScheduleDetailWorker(target_ymd, station_label, self._summary_df, self)
        self.schedule_detail_worker.done.connect(self._on_schedule_detail_done)
        self.schedule_detail_worker.error.connect(self._on_schedule_detail_error)
        self.schedule_detail_worker.start()



    def open_station_pivot_for_sn(self, sn: str, station_label: str):
        """
        Double-click on SN inside StationListDialog calls this.
        It fetches raw rows from DB only for that SN + station, then
        shows StationPivotDialog (pivot + raw).
        """
        sn = str(sn).strip()
        if not sn:
            return

        # Map station label -> SUMMARY column, then to (family, subtype)
        col_name = LABEL_TO_COL.get(station_label)
        if not col_name:
            QMessageBox.information(self, "No mapping",
                                    f"Cannot map station label '{station_label}' to SUMMARY column.")
            return


        family = subtype = None
        for cname, (fam, sub) in SUMMARY_COLUMNS:
            if cname == col_name:
                family, subtype = fam, sub
                break

        if family is None:
            QMessageBox.information(self, "No mapping",
                                    f"No family/subtype mapping for station '{station_label}'.")
            return

        try:
            conn = pyodbc.connect(DB_CONN)
        except Exception as e:
            QMessageBox.critical(self, "DB error", f"Could not connect to DB:\n{e}")
            return

        try:
            raw_df = pd.DataFrame()

            if family == "TRX":
                df_all = fetch_table_for_components_bulk(conn, TRX_TABLE, [sn])
                trx_split_raw = split_trx_types_raw(df_all)
                if subtype in trx_split_raw:
                    raw_df = trx_split_raw[subtype]
                else:
                    raw_df = df_all.copy()

            elif family == "FWWRITE":
                raw_df = fetch_table_for_components_bulk(conn, FWWRITE_TABLE, [sn])
                if isinstance(raw_df, pd.DataFrame) and not raw_df.empty:
                    raw_df = reduce_fw_latest_per_component(raw_df)
                raw_df = raw_df.copy() if isinstance(raw_df, pd.DataFrame) else pd.DataFrame()
            elif family == "MODEHOP":
                raw_df = fetch_table_for_components_bulk(conn, MODEHOP_TABLE, [sn])
                if isinstance(raw_df, pd.DataFrame) and not raw_df.empty:
                    raw_df = raw_df.copy() if isinstance(raw_df, pd.DataFrame) else pd.DataFrame()
            elif family == "BURNIN":
                raw_df = fetch_burnin_combined(conn, [sn])

            elif family in ("3TBER", "TCBER"):
                raw_df = fetch_table_for_components_bulk(conn, TABLES[family], [sn])

            elif family == "SWITCH":
                raw_df = fetch_switch_combined(conn, [sn])

            else:
                raw_df = pd.DataFrame()

            if raw_df is None or raw_df.empty:
                QMessageBox.information(self, "No data",
                                        f"No raw rows for SN {sn} at station '{station_label}'.")
                return

            # 3-temperature view for 3TBER / TCBER / Final Test
            use_three_temp = False
            if family in ("3TBER", "TCBER") or (family == "TRX" and subtype == "FINAL"):
                use_three_temp = True

            if use_three_temp:
                dlg = ThreeTempPivotDialog(sn, station_label, raw_df, parent=self)
            else:
                dlg = StationPivotDialog(sn, station_label, raw_df, parent=self)

            dlg.exec_()


        finally:
            try:
                conn.close()
            except Exception:
                pass
    



    # ── Busy helpers
    def _start_busy(self):
        self.prog.setVisible(True)
        self.prog.setValue(0)
        QApplication.setOverrideCursor(Qt.WaitCursor)
        QApplication.processEvents()

    def _set_progress(self, v):
        self.prog.setValue(max(0, min(100, v)))
        QApplication.processEvents()

    def _end_busy(self):
        self.prog.setVisible(False)
        QApplication.restoreOverrideCursor()
        QApplication.processEvents()

    # ── UI slots

    def on_source_changed(self, index):
        text = self.cmb_source.currentText()
        if text == "WO Lists":
            self._source_mode = "WO"
            self.btn_manage_lists.setEnabled(True)
            self.btn_existing_lists.setEnabled(True)
            self.btn_save_lists.setEnabled(True)
            self.chk_dr8.setVisible(False)
            self.chk_fr4.setVisible(False)
        elif text.startswith("Auto WOs"):
            self._source_mode = "DEVICE"
            self.btn_manage_lists.setEnabled(False)
            self.btn_existing_lists.setEnabled(False)
            self.btn_save_lists.setEnabled(False)
            self.chk_dr8.setVisible(True)
            self.chk_fr4.setVisible(True)
        else:
            # SN List (paste)
            self._source_mode = "SN"
            self.btn_manage_lists.setEnabled(True)
            self.btn_existing_lists.setEnabled(False)
            self.btn_save_lists.setEnabled(True)
            self.chk_dr8.setVisible(False)
            self.chk_fr4.setVisible(False)
            self.on_manage_lists()

    def on_manage_lists(self):
        dlg = ManageListsDialog(self._wo_lists, self._sn_lists, self)
        if dlg.exec_() == QDialog.Accepted:
            self._wo_lists = dlg.get_wo_lists()
            self._sn_lists = dlg.get_sn_lists()
            if self._wo_lists:
                names = ", ".join(lst["name"] for lst in self._wo_lists)
                self.lbl_lists.setText(f"Configured WO lists: {names}")
            else:
                self.lbl_lists.setText("Configured WO lists: None")
            if self._sn_lists:
                names = ", ".join(lst["name"] for lst in self._sn_lists)
                self.lbl_lists.setText(f"Configured SN lists: {names}")
            else:
                self.lbl_lists.setText("Configured SN lists: None")

    def on_existing_lists(self):
        lists = load_saved_wo_lists(SAVED_LISTS_XLSX)
        if not lists:
            QMessageBox.information(
                self,
                "No saved lists",
                "No saved WO lists found.\n\nCreate lists using 'WO Lists', then click 'Save Lists'.",
            )
            return

        dlg = WOListsDialog(lists, self)
        if dlg.exec_() == QDialog.Accepted:
            self._wo_lists = dlg.get_lists()
            if self._wo_lists:
                names = ", ".join(lst.get("name", "") for lst in self._wo_lists if lst.get("name"))
                self.lbl_lists.setText(f"Configured lists: {names}" if names else "Configured lists: (loaded)")
            else:
                self.lbl_lists.setText("Configured lists: None")

    def on_save_lists(self):
        if not self._wo_lists and not self._sn_lists:
            QMessageBox.warning(self, "No lists", "No lists to save. Use 'Manage Lists' first.")
            return
        try:
            save_wo_lists(SAVED_LISTS_XLSX, self._wo_lists)
            save_sn_lists(SAVED_LISTS_XLSX, self._sn_lists)
            QMessageBox.information(
                self,
                "Saved",
                f"WO and SN lists saved to:\n{SAVED_LISTS_XLSX}",
            )
        except Exception as e:
            QMessageBox.critical(self, "Save error", f"Could not save lists:\n{e}")

    def on_manage_fail_modes(self):
        dlg = FailModesDialog(self._fail_modes, self)
        if dlg.exec_() == QDialog.Accepted:
            self._fail_modes = dlg.get_modes()
            n = len(self._fail_modes)
            self.lbl_fail_modes.setText(f"{n} fail modes configured" if n else "0 fail modes configured")

            # Persist fail modes to the same Excel file as WO lists
            try:
                save_fail_modes(SAVED_LISTS_XLSX, self._fail_modes)
            except Exception as e:
                QMessageBox.warning(
                    self,
                    "Fail modes not saved",
                    f"Fail modes were updated, but could not be saved to Excel:\n{SAVED_LISTS_XLSX}\n\n{e}"
                )

    def on_blank_override(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("Treat blanks as PASS")
        dlg.setMinimumWidth(420)

        v = QVBoxLayout(dlg)
        v.addWidget(QLabel(
            "Select one or more stations where blank SUMMARY cells\n"
            "should be treated as PASS. This affects both Yield and Pareto.\n\n"
            "Uncheck all to clear the override."
        ))

        checks = []
        for label, _ in STATION_FLOW:
            cb = QCheckBox(label)
            cb.setChecked(label in self._blank_pass_labels)
            v.addWidget(cb)
            checks.append(cb)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        v.addWidget(buttons)

        if dlg.exec_() == QDialog.Accepted:
            selected = {cb.text() for cb in checks if cb.isChecked()}
            self._blank_pass_labels = selected

            if not selected:
                self.lbl_override.setText("Blank-pass stations: None")
            else:
                self.lbl_override.setText(
                    "Blank-pass stations: " + ", ".join(sorted(selected))
                )

            if not self._summary_df.empty:
                self._build_tabs()

    def on_bypass_clicked(self):
        """Select a station to bypass (input=output, no fails, no gating)."""
        dlg = QDialog(self)
        dlg.setWindowTitle("Bypass station")
        dlg.setMinimumWidth(420)

        v = QVBoxLayout(dlg)
        v.addWidget(QLabel(
            "Select one station to bypass.\n"
            "Bypassed station behaves as: Input=Output, Fail=0, and devices continue\n"
            "to the next station regardless of PASS/FAIL/blank at that station.\n\n"
            "Choose None to clear."
        ))

        form = QFormLayout()
        cbo = QComboBox()
        cbo.addItem("None", userData=None)
        for label, _ in STATION_FLOW:
            cbo.addItem(label, userData=label)

        current = next(iter(self._bypass_labels), None)
        if current:
            for i in range(cbo.count()):
                if cbo.itemData(i) == current:
                    cbo.setCurrentIndex(i)
                    break

        form.addRow("Bypass:", cbo)

        switch_pass_check = QCheckBox("switch test pass")
        switch_pass_check.setChecked(self._switch_test_pass)
        form.addRow(switch_pass_check)

        v.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        v.addWidget(buttons)

        if dlg.exec_() == QDialog.Accepted:
            sel = cbo.currentData()
            self._bypass_labels = {sel} if sel else set()
            self._switch_test_pass = switch_pass_check.isChecked()

            if not self._bypass_labels:
                self.lbl_bypass.setText("Bypass station: None")
            else:
                self.lbl_bypass.setText("Bypass station: " + next(iter(self._bypass_labels)))

            if not self._summary_df.empty:
                self._build_tabs()

    def on_weekly_clicked(self):
        dlg = WeeklyYieldRangeDialog(self)
        if dlg.exec_() != QDialog.Accepted:
            return

        device_type, start_ymd, end_ymd = dlg.get_values()

        # Discover weekly WO groups first (MASTER receipt date), then reuse the existing Yield pipeline (WO Lists → Get Data)
        self._start_busy()
        self.btn_get.setEnabled(False)
        self.btn_weekly.setEnabled(False)
        self.btn_schedule.setEnabled(False)

        self.weekly_worker = WeeklyWOIndexWorker(start_ymd, end_ymd, device_type, self)
        self.weekly_worker.done.connect(self._weekly_range_done)
        self.weekly_worker.error.connect(self._weekly_error)
        self.weekly_worker.start()

    def _weekly_range_done(self, wo_lists: list, title: str):
        self._end_busy()
        self.btn_get.setEnabled(True)
        self.btn_weekly.setEnabled(True)
        self.btn_schedule.setEnabled(True)

        wo_lists = wo_lists or []
        wo_lists = [lst for lst in wo_lists if lst and lst.get("wos")]
        if not wo_lists:
            QMessageBox.information(self, "Weekly Yield", "No WOs found in that date range.")
            return

        # Force WO Lists mode and load weekly WO lists (one tab per ISO week)
        self.cmb_source.setCurrentIndex(0)  # "WO Lists"
        self._wo_lists = wo_lists

        names = ", ".join(lst.get("name", "") for lst in self._wo_lists if lst.get("name"))
        self.lbl_lists.setText(f"Configured lists: {names}" if names else "Configured lists: (loaded)")

        # Reuse the normal Get Data logic
        self.on_get_clicked()

    def _weekly_error(self, msg: str):
        self._end_busy()
        self.btn_get.setEnabled(True)
        self.btn_weekly.setEnabled(True)
        self.btn_schedule.setEnabled(True)
        QMessageBox.critical(self, "Weekly Yield error", msg)



    # ── Save all tabs as Excel report
    def on_save_report_clicked(self):
        if self._summary_df is None or self._summary_df.empty:
            QMessageBox.warning(self, "No data", "Run 'Get Data' first.")
            return

        tab_lists_to_pass = []
        if self._source_mode == "WO":
            if not self._wo_lists:
                QMessageBox.warning(self, "No WO Lists", "Define WO lists first.")
                return
            tab_lists_to_pass = self._wo_lists
        elif self._source_mode == "SN":
            if not self._sn_lists:
                QMessageBox.warning(self, "No SN Lists", "Define SN lists first.")
                return
            # Adapt SN lists to expected format for ReportExportWorker
            tab_lists_to_pass = [{"name": lst["name"], "wos": lst["sns"]} for lst in self._sn_lists]
        elif self._source_mode == "DEVICE":
            if self._summary_df is None or self._summary_df.empty:
                QMessageBox.warning(self, "No data", "Run 'Get Data' first to generate Summary for Auto WOs.")
                return
            unique_wos = sorted(self._summary_df["WO"].dropna().unique().tolist())
            if not unique_wos:
                QMessageBox.warning(self, "No WOs", "No Work Orders found in current Auto WOs summary.")
                return
            tab_lists_to_pass = [{"name": "Auto WOs", "wos": unique_wos}]
        else:
            QMessageBox.warning(self, "Unsupported Mode", "Report export is not supported for the current source mode.")
            return

        if not tab_lists_to_pass:
             QMessageBox.warning(self, "No Lists", "No data lists to export for the current source mode.")
             return

        default_name = "Yield_Report.xlsx"
        out_path, _ = QFileDialog.getSaveFileName(self, "Save report", default_name, "Excel (*.xlsx)")
        if not out_path:
            return
        if not out_path.lower().endswith(".xlsx"):
            out_path += ".xlsx"

        # Start worker
        self._start_busy()
        self.btn_get.setEnabled(False)
        self.btn_weekly.setEnabled(False)
        self.btn_schedule.setEnabled(False)
        self.btn_report.setEnabled(False)

        self.report_worker = ReportExportWorker(
            out_path=out_path,
            tab_lists=tab_lists_to_pass, # Use the adapted list
            summary_df=self._summary_df,
            fail_modes=self._fail_modes,
            blank_pass_labels=self._blank_pass_labels,
            bypass_labels=self._bypass_labels,
            device_type="", # device_type is not currently captured in the main UI for reporting
            parent=self
        )
        self.report_worker.progress.connect(self._report_progress)
        self.report_worker.done.connect(self._report_done)
        self.report_worker.error.connect(self._report_error)
        self.report_worker.start()

    def _report_progress(self, pct: int, msg: str):
        self._set_progress(int(pct))
        try:
            self.statusBar().showMessage(msg)
        except Exception:
            pass

    def _report_done(self, out_path: str):
        self._end_busy()
        self.btn_get.setEnabled(True)
        self.btn_weekly.setEnabled(True)
        self.btn_schedule.setEnabled(True)
        self.btn_report.setEnabled(True)
        try:
            self.statusBar().showMessage("")
        except Exception:
            pass
        QMessageBox.information(self, "Report saved", f"Saved:\n{out_path}")

    def _report_error(self, msg: str):

            self._end_busy()

            self.btn_get.setEnabled(True)

            self.btn_weekly.setEnabled(True)

            self.btn_schedule.setEnabled(True)

            self.btn_report.setEnabled(True)

            try:

                self.statusBar().showMessage("")

            except Exception:

                pass

            QMessageBox.critical(self, "Report error", msg)

    

    def _on_schedule_detail_done(self, station_label, detail_kind, detail_df, summary_context_df):

            self._end_busy()

            dlg = StationListDialog(

                host_window=self,

                station_label=station_label,

                detail_kind=detail_kind,

                detail_df=detail_df,

                summary_context_df=summary_context_df,

                fail_modes=self._fail_modes,

                blank_pass_labels=self._blank_pass_labels,

                bypass_labels=self._bypass_labels,

                parent=self,

            )

            dlg.exec_()

    

    def _on_schedule_detail_error(self, msg: str):

            self._end_busy()

            QMessageBox.critical(self, "Schedule Detail Error", msg)

    

    def on_get_clicked(self):
        if self._source_mode == "WO":
            if not self._wo_lists:
                QMessageBox.warning(self, "No lists", "Please define at least one WO list first.")
                return
            all_wos = sorted({wo for lst in self._wo_lists for wo in lst["wos"]})
            if not all_wos:
                QMessageBox.warning(self, "No WOs", "All lists are empty.")
                return
            device_codes = []
            sn_list = []
            sn_wo_map = None
        elif self._source_mode == "SN":
            if not self._sn_lists:
                QMessageBox.warning(self, "No SN lists", "Please define at least one SN list first.")
                return

            sn_wo_map = {}
            all_sns = set()
            for lst in self._sn_lists:
                name = lst["name"]
                for sn in lst.get("sns", []):
                    sn_strip = str(sn).strip()
                    if sn_strip:
                        sn_wo_map[sn_strip] = name
                        all_sns.add(sn_strip)

            if not all_sns:
                QMessageBox.warning(self, "No SNs", "All SN lists are empty.")
                return

            all_wos = []
            device_codes = []
            sn_list = sorted(list(all_sns))
        else:
            sn_list = []
            sn_wo_map = None
            selected_devices = []
            if self.chk_dr8.isChecked():
                selected_devices.append(DR8_DEVICE_ID)
            if self.chk_fr4.isChecked():
                selected_devices.append(FR4_DEVICE_ID)
            if not selected_devices:
                QMessageBox.warning(self, "No Devices", "Please select at least DR8+ or FR4 for Auto WOs.")
                return
            all_wos = []
            device_codes = selected_devices

        use_date = self.chk_date_filter.isChecked()
        start_ymd = end_ymd = None
        if use_date:
            d1 = self.date_from.date()
            d2 = self.date_to.date()
            if d2 < d1:
                QMessageBox.warning(self, "Date range", "End date cannot be before start date.")
                return
            start_ymd = int(d1.toString("yyyyMMdd"))
            end_ymd   = int(d2.toString("yyyyMMdd"))

        if not self._fail_modes:
            ret = QMessageBox.question(
                self, "No fail modes",
                "Fail modes list is empty. You can still see yield tables, but Pareto will classify everything as OTHER.\n\n"
                "Do you want to continue?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            if ret != QMessageBox.Yes:
                return

        self._start_busy()
        self.btn_get.setEnabled(False)
        self.btn_schedule.setEnabled(False)
        self._summary_df = pd.DataFrame()
        self.tabs.clear()

        self.worker = SummaryWorker(
            self._source_mode, all_wos, device_codes,
            use_date, start_ymd, end_ymd,
            sn_list=sn_list, sn_wo_map=sn_wo_map
        )
        self.worker.progressPct.connect(self._set_progress)
        self.worker.done.connect(self._worker_done)
        self.worker.error.connect(self._worker_error)
        self.worker.start()

    def _worker_done(self, summary_df: pd.DataFrame):
        self._summary_df = summary_df.copy()
        self._ever_failed_by_station = getattr(self.worker, "ever_failed_by_station", {})
        self.btn_get.setEnabled(True)
        self.btn_schedule.setEnabled(True)
        self._end_busy()

        if summary_df.empty:
            QMessageBox.information(self, "No data", "SUMMARY is empty for these settings.")
            return

        self._build_tabs()

    def _worker_error(self, msg: str):
        self.btn_get.setEnabled(True)
        self.btn_schedule.setEnabled(True)
        self._end_busy()
        QMessageBox.critical(self, "Error", msg)

    def _build_tabs(self):
        self.tabs.clear()
        kw = dict(
            blank_pass_labels=self._blank_pass_labels,
            bypass_labels=self._bypass_labels,
            switch_test_pass=self._switch_test_pass,
        )
        fpy_kw = dict(
            blank_pass_labels=self._blank_pass_labels,
            bypass_labels=self._bypass_labels,
            ever_failed_by_station=self._ever_failed_by_station,
        )
        wo_kw = dict(
            blank_pass_labels=self._blank_pass_labels,
            bypass_labels=self._bypass_labels,
            switch_test_pass=self._switch_test_pass,
        )
        if self._source_mode == "WO":
            for lst in self._wo_lists:
                name = lst["name"]
                wos  = lst["wos"]
                funnel    = build_yield_funnel(self._summary_df, wo_subset=wos, **kw)
                fpy_df    = compute_fpy_table(self._summary_df, wo_subset=wos, **fpy_kw)
                wo_yld_df = compute_wo_yield_table(self._summary_df, wo_subset=wos, **wo_kw)
                self._add_table_tab(name, funnel, wos_list=wos, fpy_df=fpy_df, wo_yield_df=wo_yld_df)
        elif self._source_mode == "SN":
            for lst in self._sn_lists:
                name = lst["name"]
                funnel = build_yield_funnel(self._summary_df, wo_subset=[name], **kw)
                fpy_df = compute_fpy_table(self._summary_df, wo_subset=[name], **fpy_kw)
                self._add_table_tab(name, funnel, wos_list=[name], fpy_df=fpy_df, wo_yield_df=None)
        else:
            auto_label = []
            if self.chk_dr8.isChecked():
                auto_label.append("DR8+")
            if self.chk_fr4.isChecked():
                auto_label.append("FR4")
            tab_name  = "Auto – " + ("/".join(auto_label) if auto_label else "Devices")
            funnel    = build_yield_funnel(self._summary_df, wo_subset=None, **kw)
            fpy_df    = compute_fpy_table(self._summary_df, wo_subset=None, **fpy_kw)
            wo_yld_df = compute_wo_yield_table(self._summary_df, wo_subset=None, **wo_kw)
            self._add_table_tab(tab_name, funnel, wos_list=None, fpy_df=fpy_df, wo_yield_df=wo_yld_df)

        # Always add "All SNs" tab as the last tab
        self._add_sn_tab()

    def _add_sn_tab(self):
        """Add a tab showing every serial number with WO and all station pass/fail statuses."""
        if self._summary_df is None or self._summary_df.empty:
            return

        station_cols = [c for (c, _) in SUMMARY_COLUMNS if c in self._summary_df.columns]
        cols = ["WO", "SN"] + station_cols
        display_df = self._summary_df[cols].copy().reset_index(drop=True)

        model = DataFrameModel(display_df)
        tv = QTableView()
        tv.setModel(model)
        compact_table(tv, row_h=24, min_col_w=80, first_col_w=120)
        tv.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        tv.setContextMenuPolicy(Qt.CustomContextMenu)
        tv.customContextMenuRequested.connect(
            lambda pos, t=tv: self.save_table_context_menu(pos, t, "All SNs")
        )

        inner = QWidget()
        lay = QVBoxLayout(inner)
        lay.setContentsMargins(8, 8, 8, 8)
        lay.setSpacing(4)

        lbl = QLabel(f"All Serial Numbers — {len(display_df)} device(s)")
        lay.addWidget(lbl)
        lay.addWidget(tv)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(inner)

        self.tabs.addTab(scroll, "All SNs")
        self._sn_tab_view = tv

    def _add_table_tab(self, title: str, df: pd.DataFrame, wos_list,
                       fpy_df: pd.DataFrame = None, wo_yield_df: pd.DataFrame = None):
        model = DataFrameModel(df)
        tv = QTableView()
        tv.setModel(model)
        compact_table(tv, row_h=26, min_col_w=100, first_col_w=140)
        tv.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        inner = QWidget()
        lay = QVBoxLayout(inner)
        lay.setContentsMargins(8, 8, 8, 8)
        lay.setSpacing(8)

        # Header row with Pareto selector
        hdr_row = QHBoxLayout()
        hdr_row.addWidget(QLabel(f"WO List: {title}"))
        hdr_row.addSpacing(10)

        hdr_row.addWidget(QLabel("Pareto for:"))
        cmb_station = QComboBox()
        cmb_station.addItem("All stations (combined)", userData=None)
        for label, _ in STATION_FLOW:
            cmb_station.addItem(label, userData=label)
        hdr_row.addWidget(cmb_station)

        btn_refresh = QPushButton("Refresh Pareto")
        hdr_row.addWidget(btn_refresh)
        hdr_row.addStretch(1)
        lay.addLayout(hdr_row)

        # Main yield funnel table
        lay.addWidget(tv)
        self._fit_table_height_to_rows(tv)

        # First Pass Yield + WO Yield side-by-side
        if (fpy_df is not None and not fpy_df.empty) or \
           (wo_yield_df is not None and not wo_yield_df.empty):
            side_row = QHBoxLayout()
            side_row.setSpacing(16)

            if fpy_df is not None and not fpy_df.empty:
                fpy_col = QVBoxLayout()
                fpy_lbl = QLabel("First Pass Yield by Station  (double-click a number to see devices)")
                fpy_lbl.setStyleSheet("font-weight: bold; font-size: 12px; padding: 4px 2px 2px 2px;")
                fpy_col.addWidget(fpy_lbl)
                fpy_model = DataFrameModel(fpy_df)
                fpy_tv = QTableView()
                fpy_tv.setModel(fpy_model)
                compact_table(fpy_tv, row_h=24, min_col_w=90, first_col_w=180)
                fpy_tv.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
                fpy_tv.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
                self._fit_table_height_to_rows(fpy_tv)
                fpy_tv.doubleClicked.connect(
                    lambda index, t=fpy_tv, wos=wos_list: self.on_fpy_table_double_clicked(index, t, wos)
                )
                fpy_col.addWidget(fpy_tv)
                side_row.addLayout(fpy_col)

            if wo_yield_df is not None and not wo_yield_df.empty:
                wo_col = QVBoxLayout()
                wo_lbl = QLabel("Yield % by Work Order")
                wo_lbl.setStyleSheet("font-weight: bold; font-size: 12px; padding: 4px 2px 2px 2px;")
                wo_col.addWidget(wo_lbl)
                wo_model = DataFrameModel(wo_yield_df)
                wo_tv = QTableView()
                wo_tv.setModel(wo_model)
                compact_table(wo_tv, row_h=24, min_col_w=90, first_col_w=140)
                wo_tv.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
                wo_tv.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
                self._fit_table_height_to_rows(wo_tv)
                wo_col.addWidget(wo_tv)
                wo_col.addStretch()
                side_row.addLayout(wo_col)

            side_row.addStretch()
            lay.addLayout(side_row)

        # Pareto figure
        fig = Figure(figsize=(7, 3.2))
        canvas = FigureCanvas(fig)
        canvas.setMinimumHeight(430)
        lay.addWidget(canvas)

        # ── NOW create scroll and attach everything to it
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(inner)

        scroll.table_view = tv
        scroll.table_model = model
        scroll.canvas = canvas
        scroll.wo_list = list(wos_list) if wos_list else None
        scroll.title = title
        scroll.cmb_station = cmb_station

        # Signals AFTER scroll exists
        tv.doubleClicked.connect(
            lambda index, p=scroll, tab_title=title: self.on_yield_table_double_clicked(index, p, tab_title)
        )
        tv.setContextMenuPolicy(Qt.CustomContextMenu)
        tv.customContextMenuRequested.connect(
            lambda pos, t=tv, tab_title=title: self.save_table_context_menu(pos, t, tab_title)
        )

        canvas.setContextMenuPolicy(Qt.CustomContextMenu)
        canvas.customContextMenuRequested.connect(
            lambda pos, c=canvas, tab_title=title: self.save_canvas_context_menu(pos, c, tab_title)
        )

        btn_refresh.clicked.connect(lambda _, p=scroll: self.build_pareto_for_page(p, show_message=True))
        cmb_station.currentIndexChanged.connect(
            lambda _idx, p=scroll: self.build_pareto_for_page(p, show_message=False)
        )

        self.tabs.addTab(scroll, title)
        self.build_pareto_for_page(scroll, show_message=False)

    # ── Context menus

    def save_table_context_menu(self, pos, tv: QTableView, title: str):
        menu = QMenu(tv)
        act_save = menu.addAction(f"Save '{title}' table as CSV...")
        action = menu.exec_(tv.viewport().mapToGlobal(pos))
        if action == act_save:
            path, _ = QFileDialog.getSaveFileName(
                self, "Save table as CSV", f"{title}_table.csv",
                "CSV files (*.csv);;All files (*.*)"
            )
            if path:
                model = tv.model()
                if isinstance(model, DataFrameModel):
                    df = model._df
                else:
                    rows = model.rowCount()
                    cols = model.columnCount()
                    data = []
                    headers = [model.headerData(c, Qt.Horizontal) for c in range(cols)]
                    for r in range(rows):
                        row = [model.data(model.index(r, c), Qt.DisplayRole) for c in range(cols)]
                        data.append(row)
                    df = pd.DataFrame(data, columns=headers)
                try:
                    df.to_csv(path, index=False)
                except Exception as e:
                    QMessageBox.critical(self, "Save error", str(e))

    def save_canvas_context_menu(self, pos, canvas: FigureCanvas, title: str):
        menu = QMenu(canvas)
        act_save = menu.addAction(f"Save '{title}' chart as PNG...")
        action = menu.exec_(canvas.mapToGlobal(pos))
        if action == act_save:
            path, _ = QFileDialog.getSaveFileName(
                self, "Save chart as PNG", f"{title}_pareto.png",
                "PNG files (*.png);;All files (*.*)"
            )
            if path:
                try:
                    canvas.figure.savefig(path, dpi=150, bbox_inches="tight")
                except Exception as e:
                    QMessageBox.critical(self, "Save error", str(e))

    # ── Pareto
    def _fit_table_height_to_rows(self, tv: QTableView, extra=8):
        model = tv.model()
        if model is None:
            return

        rows = model.rowCount()
        if rows <= 0:
            tv.setMinimumHeight(120)
            return

        header_h = tv.horizontalHeader().height()
        row_h = tv.verticalHeader().defaultSectionSize()
        frame = tv.frameWidth() * 2

        # total height to show ALL rows (no clipping)
        total_h = header_h + (rows * row_h) + frame + extra
        tv.setMinimumHeight(total_h)

    def build_pareto_for_page(self, page: QWidget, show_message: bool = True):
        fig = page.canvas.figure
        fig.clear()

        if self._summary_df.empty:
            if show_message:
                QMessageBox.warning(self, "No data", "Run 'Get Data' first.")
            page.canvas.draw()
            return

        df = self._summary_df.copy()
        if page.wo_list:
            df = df[df["WO"].astype(str).str.strip().isin([w.strip() for w in page.wo_list])].reset_index(drop=True)

        if df.empty:
            if show_message:
                QMessageBox.information(self, "No data", "No SUMMARY rows for this tab.")
            page.canvas.draw()
            return

        cmb = getattr(page, "cmb_station", None)
        if cmb is not None:
            selected_station = cmb.currentData()  # None = combined
        else:
            selected_station = None

                # blank-as-pass labels
        blank_labels = set(self._blank_pass_labels)

        # Compute first-fail records
        records = compute_first_fail_records(df,
                                             fail_modes=self._fail_modes,
                                             blank_pass_labels=blank_labels,
                                             bypass_labels=self._bypass_labels)
        if not records:
            if show_message:
                QMessageBox.information(self, "No fails", "No devices with first-fail classified for this tab.")
            page.canvas.draw()
            return

        # 3) Build categories based on selected station
        if selected_station is None:
            # Combined: FailMode@StationTag
            labels = []
            for st, fm in records:
                if st is None or not fm:
                    continue
                station_tag = re.sub(r"[\s\-]+", "", st)
                labels.append(f"{fm}@{station_tag}")
            title_suffix = "Combined Fail Pareto"
        else:
            # Station-specific: fail modes only for that station
            labels = [
                fm for st, fm in records
                if st == selected_station and fm
            ]
            title_suffix = f"{selected_station} Fail Pareto"

        if not labels:
            if show_message:
                QMessageBox.information(self, "No fails", "No devices failed at this station with current filters.")
            page.canvas.draw()
            return

        ser = pd.Series(labels)
        counts = ser.value_counts().sort_values(ascending=False)

        # Build drill-down rows for each category (Pareto bar)
        details = compute_first_fail_details(
            df,
            self._fail_modes,
            blank_pass_labels=self._blank_pass_labels,
            bypass_labels=self._bypass_labels
        )
        cat_to_rows = {}
        if selected_station is None:
            for d in details:
                st = d.get("Station")
                fm = d.get("FailMode")
                if not st or not fm:
                    continue
                station_tag = re.sub(r"[\s\-]+", "", str(st))
                key = f"{fm}@{station_tag}"
                cat_to_rows.setdefault(key, []).append(d)
        else:
            for d in details:
                if d.get("Station") == selected_station:
                    key = d.get("FailMode")
                    cat_to_rows.setdefault(key, []).append(d)


        vals = counts.values.astype(int)
        cats = counts.index.tolist()
        cum_pct = vals.cumsum() / vals.sum() * 100.0

        ax = fig.add_subplot(111)
        x = np.arange(len(cats))

        bars = ax.bar(x, vals)

        # Make bars interactive
        for b in bars:
            try:
                b.set_picker(5)
            except Exception:
                pass
        ax.set_xlabel("Fail Mode" if selected_station is not None else "Fail Mode @ Station")
        ax.set_ylabel("Qty")

        ax.set_xticks(x)
        ax.set_xticklabels(cats, rotation=45, ha="right")

        max_val = max(vals) if len(vals) > 0 else 0
        offset = max_val * 0.05 if max_val > 0 else 0.5
        for xi, v in zip(x, vals):
            ax.text(xi, v + offset, str(v), ha="center", va="bottom", fontsize=9)

        ax2 = ax.twinx()
        ax2.plot(x, cum_pct, marker="o")
        ax2.set_ylim(0, 110)
        ax2.set_ylabel("Cumulative %")

        for xi, p in zip(x, cum_pct):
            ax2.text(xi, p + 2, f"{p:.0f}%", ha="center", va="bottom", fontsize=8)

        ax.set_title(f"{page.title} – {title_suffix}")


        # Double-click a bar to open SN list popup
        try:
            if hasattr(page, "_pareto_pick_cid") and page._pareto_pick_cid:
                page.canvas.mpl_disconnect(page._pareto_pick_cid)
        except Exception:
            pass

        def _on_pick(event):
            try:
                mouse_event = getattr(event, "mouseevent", None)
                if mouse_event is None or not mouse_event.dblclick:
                    return

                artist = event.artist
                if artist not in bars:
                    return

                # Find the index of the picked bar
                for i, bar in enumerate(bars):
                    if bar == artist:
                        cat = cats[i]
                        rows = cat_to_rows.get(cat, [])
                        title = f"{cat} – {len(rows)} devices"
                        dlg = FailModeSNDialog(title, rows, self)
                        dlg.exec_()
                        break
            except Exception:
                # never crash on UI interaction
                return

        page._pareto_pick_cid = page.canvas.mpl_connect("pick_event", _on_pick)

        fig.tight_layout()
        page.canvas.draw()

    # ── Schedule

    def on_get_schedule_clicked(self):
        qd = self.schedule_date.date()
        target_ymd = int(qd.toString("yyyyMMdd"))

        self._start_busy()
        self.btn_get.setEnabled(False)
        self.btn_schedule.setEnabled(False)

        self.schedule_worker = ScheduleWorker(target_ymd, SCHEDULE_DEVICE_IDS)
        self.schedule_worker.progressPct.connect(self._set_progress)
        self.schedule_worker.done.connect(self._schedule_done)
        self.schedule_worker.error.connect(self._schedule_error)
        self.schedule_worker.start()

    def _schedule_done(self, schedule_df: pd.DataFrame):
        self._schedule_df = schedule_df.copy()
        self.schedule_model.setDataFrame(self._schedule_df)
        self.btn_get.setEnabled(True)
        self.btn_schedule.setEnabled(True)
        self._end_busy()
        if schedule_df.empty:
            target_ymd = int(self.schedule_date.date().toString("yyyyMMdd"))
            QMessageBox.information(self, "No Schedule Data",
                f"No devices were tested on {target_ymd}.\n\n"
                "Checked all station tables directly using SID date.\n"
                "This means no records exist in the database for the selected date."
            )

    def _schedule_error(self, msg: str):
        self.btn_get.setEnabled(True)
        self.btn_schedule.setEnabled(True)
        self._end_busy()
        QMessageBox.critical(self, "Schedule error", msg)

# ───────────────────────────────────────────────────────────────
# Main entry

if __name__ == "__main__":
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    app.setFont(QFont("Segoe UI", 10))

    win = YieldFunnelWindow()
    win.show()

    sys.exit(app.exec_())