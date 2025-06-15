#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import time
import traceback

import pyodbc
import pandas as pd

from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex, QThread, pyqtSignal, QDate
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTabWidget, QGroupBox, QTableView,
    QDateEdit, QCheckBox, QMessageBox, QHeaderView,
    QScrollArea, QSplitter, QSizePolicy
)

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure


# =========================
# OPTIONAL (call BEFORE QApplication to truly control DPI)
# In your Summary main (before QApplication), you can do:
#   import slot_efficiency_mod
#   slot_efficiency_mod.set_compact_dpi_env(scale_factor="1")
# =========================
def set_compact_dpi_env(scale_factor: str = "1"):
    """
    Must be called BEFORE QApplication is created to affect Qt DPI scaling.
    - scale_factor="1" keeps UI closer to 100%
    - try "0.9" if you want even smaller
    """
    os.environ.setdefault("QT_AUTO_SCREEN_SCALE_FACTOR", "0")   # disable Qt auto scaling
    os.environ.setdefault("QT_SCALE_FACTOR", scale_factor)      # force scale
    os.environ.setdefault("QT_ENABLE_HIGHDPI_SCALING", "0")     # extra guard on some builds


# =========================
# DEFAULT CONFIG (used if Summary does not pass db_conn)
# =========================
DEFAULT_DB_CONN = (
    r"DRIVER=SQL Server;"
    r"SERVER=US_SQL01.USPL.HOME;"
    r"DATABASE=MES;"
    r"UID=LABVIEW;"
    r"PWD=LABVIEW;"
    r"Connection Timeout=10;"
)

TRX_4PART = "[MES].[dbo].[TestResult_800G_2XFR4_TRX_TEST]"
MASTER_4PART = "[MES].[dbo].[TESTRESULT_800G_MASTER]"

# Equipment -> stage mapping (your request)
STAGE_TO_EQUIP = {
    "RT": "TP2-TP3-RT-TEST",
    "LT": "800G_A007",
    "HT": "800G_A041",
}


# =========================
# Compact UI styling (works even if DPI scaling is annoying)
# =========================
def apply_compact_styles(widget: QWidget, base_pt: int = 9):
    """
    Make fonts smaller / tighter inside this tab (independent of the rest of the app).
    """
    f = widget.font() if widget.font() else QFont()
    f.setPointSize(base_pt)
    widget.setFont(f)

    # Tight, small UI
    widget.setStyleSheet("""
        QWidget { font-size: 9pt; }
        QLabel { font-size: 9pt; }
        QGroupBox { font-size: 9pt; font-weight: 600; }
        QGroupBox::title { subcontrol-origin: margin; left: 8px; padding: 0 3px 0 3px; }
        QPushButton { font-size: 9pt; padding: 4px 10px; }
        QCheckBox { font-size: 9pt; }
        QDateEdit { font-size: 9pt; padding: 2px 6px; }
        QTabBar::tab { font-size: 9pt; padding: 6px 12px; }
        QHeaderView::section { font-size: 8pt; padding: 3px; }
        QTableView { font-size: 8pt; }
    """)


# =========================
# Pandas model for QTableView
# =========================
class PandasModel(QAbstractTableModel):
    def __init__(self, df=pd.DataFrame(), parent=None):
        super().__init__(parent)
        self._df = df.copy()

    def set_df(self, df: pd.DataFrame):
        self.beginResetModel()
        self._df = df.copy()
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):
        return 0 if self._df is None else len(self._df)

    def columnCount(self, parent=QModelIndex()):
        return 0 if self._df is None else self._df.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid() or self._df is None:
            return None
        val = self._df.iat[index.row(), index.column()]
        if role == Qt.DisplayRole:
            if pd.isna(val):
                return ""
            if str(val) == "<NA>":
                return ""
            return str(val)
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if self._df is None or role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return str(self._df.columns[section])
        return str(section + 1)


# =========================
# SQL (same structure as your working query)
# =========================
def build_sql(blank_fail: bool) -> str:
    blank_rule = "1" if blank_fail else "0"

    return f"""
WITH base_latest AS (
    SELECT
        t.COMPONENTID,
        t.SID,
        t.TESTNUMBER,
        t.Location,
        t.CHNumber,
        UPPER(LTRIM(RTRIM(t.Ch_Pass_fail))) AS PF,
        CASE
            WHEN t.SID LIKE 'A%' THEN SUBSTRING(t.SID, 2, 14)
            ELSE SUBSTRING(t.SID, 1, 14)
        END AS SidTS
    FROM {TRX_4PART} t
    WHERE
        t.Location IS NOT NULL
        AND LTRIM(RTRIM(t.Location)) <> ''
        AND (
            UPPER(t.CHNumber) LIKE '%RT%' OR
            UPPER(t.CHNumber) LIKE '%LT%' OR
            UPPER(t.CHNumber) LIKE '%HT%'
        )
),
base_latest_stage AS (
    SELECT
        b.*,
        UPPER(LTRIM(RTRIM(m.equpment))) AS equpment,
        CASE
            WHEN UPPER(LTRIM(RTRIM(m.equpment))) = 'TP2-TP3-RT-TEST' THEN 'RT'
            WHEN UPPER(LTRIM(RTRIM(m.equpment))) = '800G_A007'       THEN 'LT'
            WHEN UPPER(LTRIM(RTRIM(m.equpment))) = '800G_A041'       THEN 'HT'
            ELSE NULL
        END AS Stage
    FROM base_latest b
    LEFT JOIN {MASTER_4PART} m
      ON m.TESTNUMBER = b.TESTNUMBER
),
sid_rollup AS (
    SELECT
        COMPONENTID,
        SID,
        MAX(SidTS) AS AttemptTS,
        MAX(CASE
            WHEN PF IS NULL OR PF = '' THEN {blank_rule}
            WHEN PF = 'PASS' THEN 0
            ELSE 1
        END) AS IsFail
    FROM base_latest_stage
    WHERE Stage IS NOT NULL
    GROUP BY COMPONENTID, SID
),
latest_sid AS (
    SELECT *
    FROM (
        SELECT
            s.*,
            ROW_NUMBER() OVER (
                PARTITION BY s.COMPONENTID
                ORDER BY s.AttemptTS DESC, s.SID DESC
            ) AS rn
        FROM sid_rollup s
    ) x
    WHERE x.rn = 1
),
pass_latest_sn AS (
    SELECT COMPONENTID
    FROM latest_sid
    WHERE IsFail = 0
),
base AS (
    SELECT r.*
    FROM base_latest_stage r
    JOIN pass_latest_sn p
      ON p.COMPONENTID = r.COMPONENTID
    WHERE
        r.Stage IS NOT NULL
        AND r.SidTS BETWEEN ? AND ?
),
attempts AS (
    SELECT
        COMPONENTID,
        Stage,
        SID,
        Location,
        MAX(equpment) AS equpment,
        MIN(SidTS) AS AttemptTS,
        MAX(CASE
            WHEN PF IS NULL OR PF = '' THEN {blank_rule}
            WHEN PF = 'PASS' THEN 0
            ELSE 1
        END) AS IsFail
    FROM base
    GROUP BY COMPONENTID, Stage, SID, Location
),
first_attempt AS (
    SELECT *
    FROM (
        SELECT
            a.*,
            ROW_NUMBER() OVER (
                PARTITION BY a.COMPONENTID, a.Stage
                ORDER BY a.AttemptTS ASC, a.SID ASC
            ) AS rn
        FROM attempts a
    ) x
    WHERE x.rn = 1
),
parsed_loc AS (
    SELECT
        COMPONENTID,
        Stage,
        SID,
        Location,
        AttemptTS,
        equpment,
        LEFT(LTRIM(Location), 1) AS Side,
        CASE
            WHEN CHARINDEX('Slot', Location) > 0
             AND CHARINDEX('Port', Location) > CHARINDEX('Slot', Location)
            THEN TRY_CONVERT(int,
                 LTRIM(RTRIM(
                     SUBSTRING(
                         Location,
                         CHARINDEX('Slot', Location) + 4,
                         CHARINDEX('Port', Location) - (CHARINDEX('Slot', Location) + 4)
                     )
                 ))
            )
            ELSE NULL
        END AS SlotNo,
        IsFail
    FROM first_attempt
)
SELECT
    COMPONENTID,
    Stage,
    SID,
    Location,
    AttemptTS,
    Side,
    SlotNo,
    IsFail,
    equpment
FROM parsed_loc
ORDER BY Stage, Side, SlotNo, COMPONENTID;
"""


# =========================
# DB runner with small retry
# =========================
def run_query_df(db_conn: str, start_ts: str, end_ts: str, blank_fail: bool) -> pd.DataFrame:
    sql = build_sql(blank_fail)

    retries = 2
    last = None

    for attempt in range(retries + 1):
        try:
            conn = pyodbc.connect(db_conn)
            try:
                try:
                    conn.timeout = 180
                except Exception:
                    pass

                cur = conn.cursor()
                cur.execute(sql, (start_ts, end_ts))
                rows = cur.fetchall()
                cols = [d[0] for d in cur.description] if cur.description else []
                df = pd.DataFrame.from_records(rows, columns=cols)

                if "SlotNo" in df.columns:
                    df["SlotNo"] = pd.to_numeric(df["SlotNo"], errors="coerce").astype("Int64")

                return df
            finally:
                conn.close()
        except pyodbc.Error as e:
            last = e
            msg = str(e).lower()
            if ("general network error" in msg) or ("connectionread" in msg) or ("08s01" in msg) or ("10053" in msg):
                if attempt < retries:
                    time.sleep(1.0 + attempt)
                    continue
            raise

    raise last


# =========================
# Worker thread
# =========================
class FetchWorker(QThread):
    done = pyqtSignal(pd.DataFrame)
    error = pyqtSignal(str)

    def __init__(self, db_conn: str, start_ts: str, end_ts: str, blank_fail: bool):
        super().__init__()
        self.db_conn = db_conn
        self.start_ts = start_ts
        self.end_ts = end_ts
        self.blank_fail = blank_fail

    def run(self):
        try:
            df = run_query_df(self.db_conn, self.start_ts, self.end_ts, self.blank_fail)
            self.done.emit(df)
        except Exception:
            self.error.emit(traceback.format_exc())


# =========================
# One stage panel (RT/LT/HT)
# =========================
class StagePanel(QWidget):
    def __init__(self, stage: str, parent=None):
        super().__init__(parent)
        self.stage = stage
        self.required_equipment = STAGE_TO_EQUIP[stage]
        self._df_stage = pd.DataFrame()

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        outer.addWidget(self.scroll)

        self.page = QWidget()
        self.scroll.setWidget(self.page)

        self.page_lay = QVBoxLayout(self.page)
        self.page_lay.setContentsMargins(10, 8, 10, 10)
        self.page_lay.setSpacing(10)

        self.kpi_label = QLabel("")
        self.kpi_label.setStyleSheet("font-size: 10pt; font-weight: 600;")
        self.page_lay.addWidget(self.kpi_label)

        self.a_fig = Figure(figsize=(14, 3.6))
        self.a_canvas = FigureCanvas(self.a_fig)
        self.a_canvas.setMinimumHeight(300)
        self.a_canvas.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self.b_fig = Figure(figsize=(14, 3.6))
        self.b_canvas = FigureCanvas(self.b_fig)
        self.b_canvas.setMinimumHeight(300)
        self.b_canvas.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self.a_table = QTableView()
        self.a_model = PandasModel(pd.DataFrame())
        self.a_table.setModel(self.a_model)

        self.b_table = QTableView()
        self.b_model = PandasModel(pd.DataFrame())
        self.b_table.setModel(self.b_model)

        # compact table defaults
        for tv in (self.a_table, self.b_table):
            tv.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            tv.verticalHeader().setDefaultSectionSize(18)

        self.detail_table = QTableView()
        self.detail_model = PandasModel(pd.DataFrame())
        self.detail_table.setModel(self.detail_model)
        self.detail_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.detail_table.verticalHeader().setDefaultSectionSize(18)
        self.detail_table.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.detail_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.detail_table.setMinimumHeight(520)

        a_box = QGroupBox("Side A (sorted by Fail %)")
        a_lay = QVBoxLayout()
        a_lay.addWidget(self.a_canvas)
        a_lay.addWidget(self.a_table)
        a_box.setLayout(a_lay)

        b_box = QGroupBox("Side B (sorted by Fail %)")
        b_lay = QVBoxLayout()
        b_lay.addWidget(self.b_canvas)
        b_lay.addWidget(self.b_table)
        b_box.setLayout(b_lay)

        top_split = QSplitter(Qt.Horizontal)
        top_split.addWidget(a_box)
        top_split.addWidget(b_box)
        top_split.setStretchFactor(0, 1)
        top_split.setStretchFactor(1, 1)
        top_split.setMinimumHeight(680)
        self.page_lay.addWidget(top_split)

        detail_box = QGroupBox("Devices for selected Slot (first-attempt rows within date range)")
        d_lay = QVBoxLayout()
        d_lay.addWidget(self.detail_table)
        detail_box.setLayout(d_lay)
        self.page_lay.addWidget(detail_box)

        self.page_lay.addStretch(1)

        self.a_table.clicked.connect(lambda idx: self._slot_clicked("A", idx))
        self.b_table.clicked.connect(lambda idx: self._slot_clicked("B", idx))

    def update_with_df(self, df_all: pd.DataFrame):
        if df_all is None or df_all.empty:
            self._df_stage = pd.DataFrame()
        else:
            df = df_all.copy()
            df["equpment"] = df["equpment"].astype(str).str.strip().str.upper()
            self._df_stage = df[df["equpment"] == self.required_equipment].copy()

        if not self._df_stage.empty and "SlotNo" in self._df_stage.columns:
            self._df_stage["SlotNo"] = pd.to_numeric(self._df_stage["SlotNo"], errors="coerce").astype("Int64")

        if self._df_stage.empty:
            self.kpi_label.setText(f"{self.stage} | Equipment: {self.required_equipment} | No data")
            self.a_model.set_df(pd.DataFrame())
            self.b_model.set_df(pd.DataFrame())
            self.detail_model.set_df(pd.DataFrame())
            self._plot_percent(self.a_fig, pd.DataFrame(), "Side A Fail % by Slot")
            self._plot_percent(self.b_fig, pd.DataFrame(), "Side B Fail % by Slot")
            self.scroll.verticalScrollBar().setValue(0)
            return

        total = len(self._df_stage)
        fails = int(self._df_stage["IsFail"].sum())
        rate = (fails / total * 100.0) if total else 0.0
        min_ts = self._df_stage["AttemptTS"].min()
        max_ts = self._df_stage["AttemptTS"].max()
        data_days = self._df_stage["AttemptTS"].astype(str).str[:8].nunique()

        self.kpi_label.setText(
            f"{self.stage} | Equipment: {self.required_equipment}  ||  "
            f"Total: {total}  Fail: {fails}  Fail%: {rate:.2f}  "
            f"Span: {min_ts} → {max_ts}  DataDays: {data_days}"
        )

        summary = (
            self._df_stage
            .groupby(["Side", "SlotNo"], dropna=False)
            .agg(
                Total_FirstAttempts=("COMPONENTID", "count"),
                Fail_Count=("IsFail", "sum")
            )
            .reset_index()
        )
        summary["Pass_Count"] = summary["Total_FirstAttempts"] - summary["Fail_Count"]
        summary["FailRatePct"] = (summary["Fail_Count"] / summary["Total_FirstAttempts"] * 100.0).round(2)
        summary["SlotNo"] = pd.to_numeric(summary["SlotNo"], errors="coerce").astype("Int64")

        cols = ["SlotNo", "Total_FirstAttempts", "Fail_Count", "Pass_Count", "FailRatePct"]

        a_df = summary[summary["Side"] == "A"].sort_values(
            ["FailRatePct", "Fail_Count", "Total_FirstAttempts"], ascending=[False, False, False]
        )
        b_df = summary[summary["Side"] == "B"].sort_values(
            ["FailRatePct", "Fail_Count", "Total_FirstAttempts"], ascending=[False, False, False]
        )

        self.a_model.set_df(a_df[cols].reset_index(drop=True))
        self.b_model.set_df(b_df[cols].reset_index(drop=True))

        det = self._df_stage.sort_values(
            ["IsFail", "Side", "SlotNo", "COMPONENTID"], ascending=[False, True, True, True]
        )
        self.detail_model.set_df(
            det[["Side", "SlotNo", "COMPONENTID", "SID", "Location", "IsFail", "equpment"]]
            .reset_index(drop=True)
        )

        self._plot_percent(self.a_fig, a_df, "Side A Fail % by Slot")
        self._plot_percent(self.b_fig, b_df, "Side B Fail % by Slot")

        self.scroll.verticalScrollBar().setValue(0)

    def _plot_percent(self, fig: Figure, side_df: pd.DataFrame, title: str):
        fig.clear()
        ax = fig.add_subplot(111)
        ax.set_title(title)

        if side_df is not None and not side_df.empty:
            plot_df = side_df.dropna(subset=["SlotNo"]).sort_values("SlotNo")
            x = plot_df["SlotNo"].astype(int).astype(str)
            y = plot_df["FailRatePct"].astype(float)

            ax.bar(x, y)
            ax.set_xlabel("Slot")
            ax.set_ylabel("Fail %")
            ax.set_ylim(0, 100)

        fig.tight_layout()
        fig.canvas.draw_idle()

    def _slot_clicked(self, side: str, index: QModelIndex):
        model = self.a_model if side == "A" else self.b_model
        df = model._df
        if df is None or df.empty:
            return

        slotno = df.iloc[index.row()]["SlotNo"]
        if pd.isna(slotno) or str(slotno) == "<NA>":
            return

        d = self._df_stage[
            (self._df_stage["Side"] == side) &
            (self._df_stage["SlotNo"] == slotno)
        ].copy()

        d = d.sort_values(["IsFail", "COMPONENTID"], ascending=[False, True])

        self.detail_model.set_df(
            d[["Side", "SlotNo", "COMPONENTID", "SID", "Location", "IsFail", "equpment"]]
            .reset_index(drop=True)
        )

        self.scroll.verticalScrollBar().setValue(self.scroll.verticalScrollBar().maximum())


# =========================
# THIS is what your Summary app should import/use as 4th tab
# =========================
class SlotEfficiencyTab(QWidget):
    """
    A QWidget tab wrapper around your Slot Efficiency dashboard.
    Summary app should do:
        from slot_efficiency_mod import SlotEfficiencyTab
        tab = SlotEfficiencyTab(db_conn=DB_CONN)
        main_tabs.addTab(tab, "Slot Efficiency")
    """
    def __init__(self, db_conn=None, parent=None, **kwargs):
        super().__init__(parent)

        self.db_conn = db_conn or DEFAULT_DB_CONN
        self._worker = None

        apply_compact_styles(self, base_pt=9)

        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(8)

        # Controls row
        controls = QHBoxLayout()

        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDate(QDate.currentDate().addDays(-30))

        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDate(QDate.currentDate())

        self.blank_fail = QCheckBox("Treat blank PF as FAIL")
        self.blank_fail.setChecked(False)

        self.refresh_btn = QPushButton("Refresh")
        self.refresh_btn.clicked.connect(self.refresh)

        self.status = QLabel("Ready")

        controls.addWidget(QLabel("From:"))
        controls.addWidget(self.start_date)
        controls.addWidget(QLabel("To:"))
        controls.addWidget(self.end_date)
        controls.addWidget(self.blank_fail)
        controls.addWidget(self.refresh_btn)
        controls.addStretch(1)
        controls.addWidget(self.status)

        root.addLayout(controls)

        # Inner tabs (RT/LT/HT)
        self.tabs = QTabWidget()
        self.tab_rt = StagePanel("RT")
        self.tab_lt = StagePanel("LT")
        self.tab_ht = StagePanel("HT")
        self.tabs.addTab(self.tab_rt, "RT")
        self.tabs.addTab(self.tab_lt, "LT")
        self.tabs.addTab(self.tab_ht, "HT")

        root.addWidget(self.tabs)

    def refresh(self):
        if self.start_date.date() > self.end_date.date():
            QMessageBox.warning(self, "Date Range", "From date cannot be after To date.")
            return

        start_ts = self.start_date.date().toString("yyyyMMdd") + "000000"
        end_ts = self.end_date.date().toString("yyyyMMdd") + "235959"

        self.status.setText("Fetching…")
        self.refresh_btn.setEnabled(False)

        self._worker = FetchWorker(self.db_conn, start_ts, end_ts, self.blank_fail.isChecked())
        self._worker.done.connect(self.on_data)
        self._worker.error.connect(self.on_err)
        self._worker.start()

    def on_data(self, df: pd.DataFrame):
        self.status.setText(f"Loaded rows: {len(df)}")
        self.refresh_btn.setEnabled(True)

        self.tab_rt.update_with_df(df)
        self.tab_lt.update_with_df(df)
        self.tab_ht.update_with_df(df)

    def on_err(self, msg: str):
        self.status.setText("Error")
        self.refresh_btn.setEnabled(True)
        QMessageBox.critical(self, "DB Error", msg)


# =========================
# Optional standalone test
# =========================
if __name__ == "__main__":
    # If you want smaller UI when running standalone:
    # set_compact_dpi_env("1")  # or "0.9"

    app = QApplication([])
    w = SlotEfficiencyTab()
    w.setWindowTitle("Slot Efficiency (Tab Test)")
    w.resize(1600, 900)
    w.show()
    app.exec_()
