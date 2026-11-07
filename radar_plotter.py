import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
import matplotlib.dates as mdates
import numpy as np
import os
import io
import math
import re
import threading
import urllib.request
import csv
from copy import copy
from datetime import datetime

try:
    from openpyxl import load_workbook
    from openpyxl.cell.cell import MergedCell
    from openpyxl.styles import PatternFill
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False

try:
    from PIL import Image, ImageTk
    _PIL_OK = True
except ImportError:
    _PIL_OK = False


# ── CrosshairCursor – 1 tengely (Tab 1 & 2) ──────────────────────────────────
class CrosshairCursor:
    """Mozgatható kereszthajszál-kurzor egy matplotlib tengelyhez."""

    def __init__(self, ax, canvas, coord_var, label_prefix="", x_is_date=False, hover_lookup=None):
        self.ax = ax
        self.canvas = canvas
        self.coord_var = coord_var
        self.label_prefix = label_prefix
        self.x_is_date = x_is_date
        self.hover_lookup = hover_lookup
        self.enabled = False

        self._hline = ax.axhline(ax.get_ylim()[0], color="#e00000", linewidth=0.9,
                                 linestyle="--", alpha=0.8, visible=False)
        self._hlines = [ax.axhline(ax.get_ylim()[0], color="#0078d4", linewidth=0.9,
                                   linestyle="--", alpha=0.8, visible=False)
                        for _ in range(4)]
        self._vline = ax.axvline(ax.get_xlim()[0], color="#e00000", linewidth=0.9,
                                 linestyle="--", alpha=0.8, visible=False)
        self._last_mouse_xy = None
        self._text = ax.text(0.01, 0.97, "", transform=ax.transAxes,
                             fontsize=9, va="top", ha="left",
                             bbox=dict(boxstyle="round,pad=0.3", fc="lightyellow",
                                       ec="#aaa", alpha=0.85),
                             visible=False, zorder=10)
        self._cid = canvas.mpl_connect("motion_notify_event", self._on_move)

    def _ensure_hline_count(self, count):
        while len(self._hlines) < count:
            self._hlines.append(
                self.ax.axhline(
                    self.ax.get_ylim()[0], color="#0078d4", linewidth=0.9,
                    linestyle="--", alpha=0.8, visible=False
                )
            )

    def _safe_interp(self, x, line):
        x_data = np.asarray(line.get_xdata())
        y_data = np.asarray(line.get_ydata())
        if x_data.size < 2 or y_data.size < 2:
            return np.nan

        if x_data.dtype == object:
            try:
                x_data = np.array([
                    mdates.date2num(val) if isinstance(val, datetime) else float(val)
                    for val in x_data
                ], dtype=float)
            except Exception:
                x_data = np.array([float(val) if val is not None else np.nan for val in x_data], dtype=float)
        elif np.issubdtype(x_data.dtype, np.datetime64):
            x_data = mdates.date2num(x_data)

        if y_data.dtype == object:
            y_data = np.array([float(val) if val is not None else np.nan for val in y_data], dtype=float)
        else:
            y_data = y_data.astype(float, copy=False)

        mask = ~np.isnan(x_data) & ~np.isnan(y_data)
        if mask.sum() < 2:
            return np.nan
        x_clean = x_data[mask]
        y_clean = y_data[mask]
        order = np.argsort(x_clean)
        x_sorted = x_clean[order]
        y_sorted = y_clean[order]
        return float(np.interp(x, x_sorted, y_sorted, left=np.nan, right=np.nan))

    def _on_move(self, event):
        if not self.enabled or event.inaxes is not self.ax or event.xdata is None:
            self._set_visible(False)
            self.canvas.draw_idle()
            return

        if event.x is not None and event.y is not None and self._last_mouse_xy is not None:
            dx = abs(event.x - self._last_mouse_xy[0])
            dy = abs(event.y - self._last_mouse_xy[1])
            if dx < 3 and dy < 3:
                return
        self._last_mouse_xy = (event.x, event.y)

        x, y = event.xdata, event.ydata
        self._vline.set_xdata([x, x])
        self._vline.set_visible(True)

        visible_data_lines = [
            line for line in self.ax.lines
            if not line.get_label().startswith("_") and line.get_visible()
        ]
        self._ensure_hline_count(len(visible_data_lines))

        line_labels = []
        active = 0
        for line in visible_data_lines:
            label = line.get_label()
            y_val = self._safe_interp(x, line)
            if np.isnan(y_val):
                continue
            hline = self._hlines[active]
            hline.set_ydata([y_val, y_val])
            hline.set_color(line.get_color() or "#0078d4")
            hline.set_visible(True)
            line_labels.append(f"{label}: {y_val:.4f}")
            active += 1

        for hidden in self._hlines[active:]:
            hidden.set_visible(False)

        if self.x_is_date and x is not None:
            try:
                x_label = mdates.num2date(x).strftime("%Y-%m-%d %H:%M")
                x_label_name = "Dátum"
            except Exception:
                x_label = f"{x:.3f}"
                x_label_name = "Mag"
        else:
            x_label = f"{x:.3f} km"
            x_label_name = "Mag"

        if line_labels:
            label = f"{self.label_prefix}  {x_label_name}: {x_label}   " + "   ".join(line_labels)
            self._hline.set_visible(False)
        else:
            self._hline.set_ydata([y, y])
            self._hline.set_visible(True)
            label = f"{self.label_prefix}  {x_label_name}: {x_label}   Y: {y:.4f}"

        if self.hover_lookup is not None and x is not None and y is not None:
            try:
                hover_text = self.hover_lookup(x, y)
                if hover_text:
                    label = f"{label}   {hover_text}"
            except Exception:
                pass

        self._text.set_text(label)
        self._text.set_visible(True)
        self.coord_var.set(label.strip())
        self.canvas.draw_idle()

    def _set_visible(self, state):
        self._hline.set_visible(state)
        self._vline.set_visible(state)
        self._text.set_visible(state)
        for hline in self._hlines:
            hline.set_visible(state)
        if not state:
            self._last_mouse_xy = None
            self.coord_var.set("")

    def toggle(self, state: bool):
        self.enabled = state
        if not state:
            self._set_visible(False)
            self.canvas.draw_idle()


# ── ComparisonCrosshair – 2 szinkronizált tengely (Tab 3) ─────────────────────
def _interp_safe(x, h_arr, y_arr):
    """NaN-mentes lineáris interpoláció."""
    mask = ~np.isnan(y_arr)
    if mask.sum() < 2:
        return np.nan
    return float(np.interp(x, h_arr[mask], y_arr[mask]))


class ComparisonCrosshair:
    """Szinkronizált kurzor az összehasonlító grafikonokhoz.

    - Függőleges vonal (piros): mindkét tengelyen azonos magassági pozíción
    - Vízszintes vonalak: a két tengely aktív adatsorai
    - Szövegcímkék: mindkét tengelyen kiírja az aktuális értékeket
    """

    def __init__(self, ax1, ax2, canvas, coord_var):
        self.ax1 = ax1
        self.ax2 = ax2
        self.canvas = canvas
        self.coord_var = coord_var
        self.enabled = False
        self._hline_cache1 = []
        self._hline_cache2 = []

        # Vonalak – ax1
        self._v1        = ax1.axvline(0, color="#cc0000", lw=1.0, ls="--", alpha=0.85, visible=False)
        self._txt1 = ax1.text(0.01, 0.97, "", transform=ax1.transAxes,
                              fontsize=8.5, va="top", ha="left",
                              bbox=dict(boxstyle="round,pad=0.3", fc="lightyellow",
                                        ec="#aaa", alpha=0.88),
                              visible=False, zorder=10)
        # Vonalak – ax2
        self._v2        = ax2.axvline(0, color="#cc0000", lw=1.0, ls="--", alpha=0.85, visible=False)
        self._txt2 = ax2.text(0.01, 0.97, "", transform=ax2.transAxes,
                              fontsize=8.5, va="top", ha="left",
                              bbox=dict(boxstyle="round,pad=0.3", fc="lightyellow",
                                        ec="#aaa", alpha=0.88),
                              visible=False, zorder=10)

        self._cid = canvas.mpl_connect("motion_notify_event", self._on_move)

    def _visible_series_lines(self, ax):
        return [line for line in ax.lines if not line.get_label().startswith("_") and line.get_visible()]

    def _ensure_hlines(self, cache, ax, count):
        while len(cache) < count:
            cache.append(ax.axhline(0, color="#0078d4", lw=0.9, ls="--", alpha=0.8, visible=False))

    def _interp_line(self, x, line):
        x_data = np.asarray(line.get_xdata())
        y_data = np.asarray(line.get_ydata())
        if x_data.size < 2 or y_data.size < 2:
            return np.nan
        if y_data.dtype == object:
            try:
                y_data = np.array([float(val) if val is not None else np.nan for val in y_data], dtype=float)
            except Exception:
                return np.nan
        else:
            y_data = y_data.astype(float, copy=False)
        x_data = x_data.astype(float, copy=False)
        mask = ~np.isnan(x_data) & ~np.isnan(y_data)
        if mask.sum() < 2:
            return np.nan
        x_clean = x_data[mask]
        y_clean = y_data[mask]
        order = np.argsort(x_clean)
        return float(np.interp(x, x_clean[order], y_clean[order], left=np.nan, right=np.nan))

    def _on_move(self, event):
        if not self.enabled or event.inaxes not in (self.ax1, self.ax2):
            self._set_all_visible(False)
            self.canvas.draw_idle()
            return

        x = event.xdata  # magasság

        for vl in (self._v1, self._v2):
            vl.set_xdata([x, x])
            vl.set_visible(True)

        # 1. grafikon
        lbl1 = f"Mag: {x:.3f} km"
        active1 = self._visible_series_lines(self.ax1)
        self._ensure_hlines(self._hline_cache1, self.ax1, len(active1))
        for line, hline in zip(active1, self._hline_cache1):
            y_val = self._interp_line(x, line)
            if np.isnan(y_val):
                hline.set_visible(False)
                continue
            hline.set_ydata([y_val, y_val])
            hline.set_color(line.get_color() or "#0078d4")
            hline.set_visible(True)
            lbl1 += f"   {line.get_label()}: {y_val:.4f}"
        for hline in self._hline_cache1[len(active1):]:
            hline.set_visible(False)
        self._txt1.set_text(lbl1)
        self._txt1.set_visible(True)

        # 2. grafikon
        lbl2 = f"Mag: {x:.3f} km"
        active2 = self._visible_series_lines(self.ax2)
        self._ensure_hlines(self._hline_cache2, self.ax2, len(active2))
        for line, hline in zip(active2, self._hline_cache2):
            y_val = self._interp_line(x, line)
            if np.isnan(y_val):
                hline.set_visible(False)
                continue
            hline.set_ydata([y_val, y_val])
            hline.set_color(line.get_color() or "#0078d4")
            hline.set_visible(True)
            lbl2 += f"   {line.get_label()}: {y_val:.4f}"
        for hline in self._hline_cache2[len(active2):]:
            hline.set_visible(False)
        self._txt2.set_text(lbl2)
        self._txt2.set_visible(True)

        self.coord_var.set(f"1. fájl → {lbl1}     ||     2. fájl → {lbl2}")
        self.canvas.draw_idle()

    def _set_all_visible(self, state):
        objs = [self._v1, self._v2, self._txt1, self._txt2]
        objs.extend(self._hline_cache1)
        objs.extend(self._hline_cache2)
        for obj in objs:
            obj.set_visible(state)
        if not state:
            self.coord_var.set("")

    def toggle(self, state: bool):
        self.enabled = state
        if not state:
            self._set_all_visible(False)
            self.canvas.draw_idle()


# ── Állomás / dátum visszafejtése fájlnévből ─────────────────────────────────
_STATIONS = {
    "BUD": "Budapest",
    "HAR": "Hármashegy",
    "NAP": "Napkor",
    "POG": "Pogányvár",
    "SZE": "Szentes",
}


def decode_filename(filename):
    name = os.path.splitext(filename)[0]
    if len(name) < 15:
        return None
    code = name[:3].upper()
    digits = name[3:15]
    station = _STATIONS.get(code)
    if station is None or not digits.isdigit():
        return None
    try:
        return (f"{station}  |  {digits[0:4]}. {digits[4:6]}. {digits[6:8]}."
                f"  {digits[8:10]}:{digits[10:12]}")
    except Exception:
        return None


def build_radar_url(filename):
    name = os.path.splitext(filename)[0]
    if len(name) < 15:
        return None, None
    code = name[:3].upper()
    digits = name[3:15]
    if code not in _STATIONS or not digits.isdigit():
        return None, None
    try:
        y, mo, d = digits[0:4], digits[4:6], digits[6:8]
        h = digits[8:10]
        mi = (int(digits[10:12]) // 5) * 5
        img_url  = (f"https://www.metnet.hu/img/radar_metnet/"
                    f"{y}/{mo}/{d}/composite_{y}{mo}{d}_{h}{mi:02d}.jpg")
        page_url = (f"https://www.metnet.hu/radar/kompozit/"
                    f"{y}/{mo}/{d}/{h}:{mi:02d}/orszagos")
        return img_url, page_url
    except Exception:
        return None, None


def parse_file(filepath):
    heights, zdr, zatlag = [], [], []
    with open(filepath, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            parts = line.split()
            if len(parts) < 3:
                continue
            try:
                if parts[0].upper() == "NAN":
                    continue
                h  = float(parts[0])
                z  = np.nan if parts[1].upper() == "NAN" else float(parts[1])
                za = np.nan if parts[2].upper() == "NAN" else float(parts[2])
                heights.append(h)
                zdr.append(z)
                zatlag.append(za)
            except ValueError:
                continue
    return np.array(heights), np.array(zdr), np.array(zatlag)


def parse_file_tab2(filepath):
    """A 2. fülhöz tartozó, legfeljebb 12 oszlopos beolvasás.

    Elvárt oszlopsorrend:
    Magasság, ZDR, ZH, ZV, UH, UV, P, RHOHV, SNRV, SNRH, CCORH, CCORV
    """
    values_3 = {"ZDR": [], "Zátlag": []}
    col_names_12 = [
        "ZDR", "ZH", "ZV", "UH", "UV", "P", "RHOHV",
        "SNRV", "SNRH", "CCORH", "CCORV"
    ]
    values_12 = {name: [] for name in col_names_12}
    heights = []
    mode = None

    def parse_token(parts, index):
        if index >= len(parts):
            return np.nan
        token = parts[index]
        if token.upper() == "NAN":
            return np.nan
        try:
            return float(token)
        except ValueError:
            return np.nan

    with open(filepath, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            parts = line.split()
            if len(parts) < 3:
                continue

            try:
                if parts[0].upper() == "NAN":
                    continue
                h = float(parts[0])
            except ValueError:
                continue

            if mode is None:
                mode = "12" if len(parts) >= 12 else "3"

            heights.append(h)
            if mode == "3":
                values_3["ZDR"].append(parse_token(parts, 1))
                values_3["Zátlag"].append(parse_token(parts, 2))
            else:
                values_12["ZDR"].append(parse_token(parts, 1))
                values_12["ZH"].append(parse_token(parts, 2))
                values_12["ZV"].append(parse_token(parts, 3))
                values_12["UH"].append(parse_token(parts, 4))
                values_12["UV"].append(parse_token(parts, 5))
                values_12["P"].append(parse_token(parts, 6))
                values_12["RHOHV"].append(parse_token(parts, 7))
                values_12["SNRV"].append(parse_token(parts, 8))
                values_12["SNRH"].append(parse_token(parts, 9))
                values_12["CCORH"].append(parse_token(parts, 10))
                values_12["CCORV"].append(parse_token(parts, 11))

    if mode == "3":
        return np.array(heights), {k: np.array(v) for k, v in values_3.items()}, mode
    return np.array(heights), {k: np.array(v) for k, v in values_12.items()}, "12"


# ── Főablak ───────────────────────────────────────────────────────────────────
class RadarPlotterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Radar Adatok Vizualizátora")
        self.geometry("1150x840")
        self.resizable(True, True)
        self.configure(bg="#f0f0f0")

        # 1. fájl
        self.filepath = tk.StringVar(value="Nincs fájl betöltve")
        self._fullpath = None
        self.heights = self.zdr = self.zatlag = None
        self.tab2_heights = np.array([])
        self.tab2_series = {}
        self._tab2_mode = None
        self._tab2_series_order = []
        self._tab2_series_orders = {
            "3": ["ZDR", "Zátlag"],
            "12": ["ZDR", "ZH", "ZV", "UH", "UV", "P", "RHOHV",
                   "SNRV", "SNRH", "CCORH", "CCORV"],
        }
        self._tab2_defaults = {
            "3": {"ZDR": True, "Zátlag": True},
            "12": {"ZDR": True, "ZH": False, "ZV": False, "UH": False,
                   "UV": False, "P": False, "RHOHV": False, "SNRV": False,
                   "SNRH": False, "CCORH": False, "CCORV": False},
        }
        self._tab2_visible_vars = {}
        self._tab2_palette = {
            "ZDR": "#0078d4",
            "Zátlag": "#e05c00",
            "ZH": "#e05c00",
            "ZV": "#2e8b57",
            "UH": "#8a2be2",
            "UV": "#d62728",
            "P": "#17becf",
            "RHOHV": "#bcbd22",
            "SNRV": "#ff1493",
            "SNRH": "#6b8e23",
            "CCORH": "#8b4513",
            "CCORV": "#1f77b4",
        }
        self._tab2_markers = {
            "ZDR": "o",
            "Zátlag": "s",
            "ZH": "s",
            "ZV": "^",
            "UH": "D",
            "UV": "v",
            "P": "P",
            "RHOHV": "X",
            "SNRV": "<",
            "SNRH": ">",
            "CCORH": "h",
            "CCORV": "*",
        }

        self.cmp_heights1 = np.array([])
        self.cmp_series1 = {}
        self.cmp_mode1 = None
        self.cmp_heights2 = np.array([])
        self.cmp_series2 = {}
        self.cmp_mode2 = None
        self._cmp_series_order = []
        self._cmp_visible_vars = {}
        self._cmp_palette = self._tab2_palette.copy()
        self._cmp_markers = self._tab2_markers.copy()
        self._cmp_controls_frame = None
        self._cmp_controls_strip = None

        # 2. fájl (összehasonlításhoz)
        self.heights2 = self.zdr2 = self.zatlag2 = None
        self._filepath2_var = tk.StringVar(value="Nincs 2. fájl betöltve")

        # Kurzor – Tab 1 & 2
        self._crosshair_on = tk.BooleanVar(value=False)
        self._cursors = []
        self._coord_var = tk.StringVar(value="")

        # Kurzor – Tab 3
        self._cmp_cursor_on = tk.BooleanVar(value=False)
        self._cmp_cursor = None
        self._cmp_coord_var = tk.StringVar(value="")

        # Radarkép – 1. állomás
        self._radar_url = None
        self._radar_image_ref = None
        
        # Radarkép – 2. állomás
        self._radar_url2 = None
        self._radar_image_ref2 = None

        # Havi mappa kiválasztása
        self._month_folder = None
        self._month_folder_var = tk.StringVar(value="Nincs hónap mappa kiválasztva")
        self._month_status_var = tk.StringVar(value="Válasszon ki egy hónap mappát a ZDR pontfelhőhöz.")
        self._month_daily_avg_enabled = tk.BooleanVar(value=False)
        self._month_ci_enabled = tk.BooleanVar(value=False)
        self._month_break_enabled = tk.BooleanVar(value=False)
        self._month_break_count_var = tk.IntVar(value=1)
        self._month_break_mean_mode = tk.BooleanVar(value=False)
        self._month_break_points = []
        self._month_break_drag_index = None
        self._month_break_x = None
        self._month_break_dragging = False
        self._month_break_cid_press = None
        self._month_break_cid_move = None
        self._month_break_cid_release = None
        self._month_break_xlim = None
        self._month_break_ax = None
        self._month_cluster_x = np.array([])
        self._month_cluster_y = np.array([])
        self._month_dates = np.array([])
        self._month_zdr_values = np.array([])
        self._month_zdr_samples = []
        self._height_folder = None
        self._height_folder_var = tk.StringVar(value="Nincs hónap mappa kiválasztva")
        self._height_status_var = tk.StringVar(value="Válasszon ki egy hónap mappát a magasságpontfelhőhöz.")
        self._height_filter_enabled = tk.BooleanVar(value=False)
        self._height_filter_low_var = tk.DoubleVar(value=-10.0)
        self._height_filter_high_var = tk.DoubleVar(value=10.0)
        self._height_stats_enabled = tk.BooleanVar(value=False)
        self._height_stats_var = tk.StringVar(value="Intervallum statisztika kikapcsolva.")
        self._height_stats_range = None
        self._height_stats_cid = None
        self._height_cloud_ax1 = None

        # Kalibralas elotti/utani tablazatexport
        self._table_before_folder = None
        self._table_after_folder = None
        self._table_template_path = os.path.join(os.path.dirname(__file__), "tablazat.xlsx")
        self._table_before_folder_var = tk.StringVar(value="Nincs mappa kiválasztva")
        self._table_after_folder_var = tk.StringVar(value="Nincs mappa kiválasztva")
        self._table_export_status_var = tk.StringVar(value="Válassza ki mindkét kalibrálási mappát.")

        # Data Quality Control tab
        self._dqc_folder = None
        self._dqc_folder_var = tk.StringVar(value="Nincs Data Quality mappa kiválasztva")
        self._dqc_status_var = tk.StringVar(value="Válasszon ki egy Data Quality mappát.")
        self._dqc_height_low_var = tk.DoubleVar(value=0.0)
        self._dqc_height_high_var = tk.DoubleVar(value=10.0)
        self._dqc_zhzv_low_var = tk.DoubleVar(value=0.0)
        self._dqc_zhzv_high_var = tk.DoubleVar(value=30.0)
        self._dqc_rhohv_min_var = tk.DoubleVar(value=0.98)
        self._dqc_snr_min_var = tk.DoubleVar(value=30.0)
        self._dqc_files = []

        self._melt_folder = None
        self._melt_folder_var = tk.StringVar(value="Nincs év mappa kiválasztva")
        self._melt_status_var = tk.StringVar(value="Válasszon ki egy éves mappastruktúrát az olvadási réteg értékeihez.")
        self._melt_summary_var = tk.StringVar(value="")
        self._melt_monthly_values = np.full(12, np.nan, dtype=float)
        self._melt_filter_enabled = tk.BooleanVar(value=False)
        self._melt_filter_low_var = tk.DoubleVar(value=140.0)
        self._melt_filter_high_var = tk.DoubleVar(value=10000.0)

        # Δ0-hoz tab magasságküszöb
        self._delta_limit_var = tk.DoubleVar(value=0.0)
        self._delta_limit_label = tk.StringVar(value="Felső határ: – km")
        self._delta_limit_scale = None

        self._build_ui()

    # ── UI felépítés ──────────────────────────────────────────────────────────
    def _build_ui(self):
        ctrl = tk.Frame(self, bg="#f0f0f0", pady=8, padx=10)
        ctrl.pack(fill=tk.X)

        tk.Button(ctrl, text="📂  Fájl megnyitása (.txt)",
                  command=self.open_file,
                  bg="#0078d4", fg="white", font=("Segoe UI", 10, "bold"),
                  relief=tk.FLAT, padx=12, pady=6, cursor="hand2",
                  ).pack(side=tk.LEFT)

        tk.Label(ctrl, textvariable=self.filepath,
                 bg="#f0f0f0", fg="#444", font=("Segoe UI", 9),
                 ).pack(side=tk.LEFT, padx=14)

        self._cursor_btn = tk.Checkbutton(
            ctrl, text="✛  Kurzor",
            variable=self._crosshair_on, command=self._toggle_crosshair,
            bg="#f0f0f0", fg="#222", selectcolor="#d0eaff",
            activebackground="#f0f0f0", font=("Segoe UI", 10),
            relief=tk.FLAT, padx=10, pady=4, cursor="hand2", indicatoron=False, bd=1,
        )
        self._cursor_btn.pack(side=tk.RIGHT, padx=6)

        tk.Label(ctrl, textvariable=self._coord_var,
                 bg="#f0f0f0", fg="#b00020", font=("Consolas", 9), anchor="e",
                 ).pack(side=tk.RIGHT, padx=10)

        # Notebook – 4 tab
        style = ttk.Style(self)
        style.configure("TNotebook.Tab", font=("Segoe UI", 10))
        nb = ttk.Notebook(self)
        nb.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 8))

        self.tab2 = tk.Frame(nb, bg="white")
        nb.add(self.tab2, text="  Magasság – ZDR & Zátlag  ")

        self.tab3 = tk.Frame(nb, bg="white")
        nb.add(self.tab3, text="  ⚖  Összehasonlítás  ")

        self.tab5 = tk.Frame(nb, bg="white")
        nb.add(self.tab5, text="  ΔZDR & ΔZátlag  ")

        # New tab: differences relative to zero based on measurement dates
        self.tab6 = tk.Frame(nb, bg="white")
        nb.add(self.tab6, text="  Δ0-hoz (időbeli változás)  ")

        self.tab4 = tk.Frame(nb, bg="#1a1a2e")
        nb.add(self.tab4, text="  🛰  Radarkép (1. állomás)  ")
        
        self.tab7 = tk.Frame(nb, bg="#1a1a2e")
        nb.add(self.tab7, text="  🛰  Radarkép (2. állomás)  ")

        self.tab8 = tk.Frame(nb, bg="white")
        nb.add(self.tab8, text="  Havi ZDR pontfelhő  ")

        self.tab9 = tk.Frame(nb, bg="white")
        nb.add(self.tab9, text="  Magasság alapú pontfelhő  ")

        self.tab12 = tk.Frame(nb, bg="white")
        nb.add(self.tab12, text="  ZDR táblázat export  ")

        self.tab10 = tk.Frame(nb, bg="white")
        nb.add(self.tab10, text="  Data Quality Control  ")

        self.tab11 = tk.Frame(nb, bg="white")
        nb.add(self.tab11, text="  Olvadási réteg magasságok  ")

        self._init_figure(self.tab2, "tab2")
        self._init_figure(self.tab5, "tab5")
        self._init_tab6_controls(self.tab6)
        self._init_figure(self.tab6, "tab6")
        self._init_tab3_comparison(self.tab3)
        self._init_tab4_radar(self.tab4)
        self._init_tab7_radar(self.tab7)
        self._init_tab8_month(self.tab8)
        self._init_figure(self.tab8, "tab8")
        self._init_tab9_height(self.tab9)
        self._init_tab12_table_export(self.tab12)
        self._init_tab10_dqc(self.tab10)
        self._init_tab11_melt_layer(self.tab11)

        # Alsó infó sávok
        self.file_info = tk.StringVar(value="")
        tk.Label(self, textvariable=self.file_info,
                 bg="#d6eaf8", fg="#1a5276", font=("Segoe UI", 9, "bold"),
                 anchor="w", padx=10, pady=2,
                 ).pack(fill=tk.X, side=tk.BOTTOM)

        self.status = tk.StringVar(value="Nyisson meg egy .txt fájlt a megjelenítéshez.")
        tk.Label(self, textvariable=self.status,
                 bg="#e0e0e0", fg="#333", font=("Segoe UI", 9),
                 anchor="w", padx=10,
                 ).pack(fill=tk.X, side=tk.BOTTOM)

    # ── Tab 1 & 2 figure init ─────────────────────────────────────────────────
    def _init_figure(self, parent, tag):
        fig = Figure(figsize=(10, 6), dpi=96, facecolor="white")
        canvas = FigureCanvasTkAgg(fig, master=parent)
        canvas.draw()

        if tag == "tab2":
            cbar = tk.Frame(parent, bg="#f7f7f7", pady=6, padx=10, relief=tk.GROOVE, bd=1)
            cbar.pack(fill=tk.X)
            self._tab2_controls_frame = cbar
            self._tab2_controls_label = tk.Label(
                cbar, text="Adatsorok:", bg="#f7f7f7", fg="#333",
                font=("Segoe UI", 10, "bold")
            )
            self._tab2_controls_label.pack(side=tk.LEFT, padx=(0, 8))
            self._tab2_controls_strip = tk.Frame(cbar, bg="#f7f7f7")
            self._tab2_controls_strip.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self._render_tab2_controls()

        tb = tk.Frame(parent, bg="white")
        tb.pack(fill=tk.X)
        NavigationToolbar2Tk(canvas, tb).update()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        if tag == "tab2":
            self.fig2, self.canvas2 = fig, canvas
        elif tag == "tab5":
            self.fig3, self.canvas3 = fig, canvas
        elif tag == "tab6":
            self.fig4, self.canvas4 = fig, canvas
        elif tag == "tab8":
            self.fig5, self.canvas5 = fig, canvas
        elif tag == "tab9":
            self.fig6, self.canvas6 = fig, canvas
        else:
            self.fig2, self.canvas2 = fig, canvas

    # ── Tab 4/5 – Különbségek ábrázolása ───────────────────────────────────────
    def _init_tab3_comparison(self, parent):
        # Vezérlő sáv a 2. fájlhoz
        cbar = tk.Frame(parent, bg="#f0f7f0", pady=6, padx=10, relief=tk.GROOVE, bd=1)
        cbar.pack(fill=tk.X)

        tk.Button(cbar, text="📂  2. állomás betöltése (.txt)",
                  command=self.open_file2,
                  bg="#217a3c", fg="white", font=("Segoe UI", 10, "bold"),
                  relief=tk.FLAT, padx=10, pady=5, cursor="hand2",
                  ).pack(side=tk.LEFT)

        tk.Label(cbar, textvariable=self._filepath2_var,
                 bg="#f0f7f0", fg="#444", font=("Segoe UI", 9),
                 ).pack(side=tk.LEFT, padx=12)

        self._cmp_cursor_btn = tk.Checkbutton(
            cbar, text="✛  Kurzor",
            variable=self._cmp_cursor_on, command=self._toggle_cmp_cursor,
            bg="#f0f7f0", fg="#222", selectcolor="#c8e6c9",
            activebackground="#f0f7f0", font=("Segoe UI", 10),
            relief=tk.FLAT, padx=10, pady=4, cursor="hand2", indicatoron=False, bd=1,
        )
        self._cmp_cursor_btn.pack(side=tk.RIGHT, padx=6)

        tk.Label(cbar, textvariable=self._cmp_coord_var,
                 bg="#f0f7f0", fg="#b00020", font=("Consolas", 8), anchor="e",
                 ).pack(side=tk.RIGHT, padx=8)

        self._cmp_controls_frame = tk.Frame(parent, bg="#f7f7f7", pady=6, padx=10, relief=tk.GROOVE, bd=1)
        self._cmp_controls_frame.pack(fill=tk.X)
        tk.Label(self._cmp_controls_frame, text="Adatsorok:", bg="#f7f7f7", fg="#333",
             font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(0, 8))
        self._cmp_controls_strip = tk.Frame(self._cmp_controls_frame, bg="#f7f7f7")
        self._cmp_controls_strip.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Matplotlib figure – 2 subplot egymás alatt
        self.figcmp = Figure(figsize=(10, 8), dpi=96, facecolor="white")
        self.canvascmp = FigureCanvasTkAgg(self.figcmp, master=parent)
        self.canvascmp.draw()
        tb = tk.Frame(parent, bg="white")
        tb.pack(fill=tk.X)
        NavigationToolbar2Tk(self.canvascmp, tb).update()
        self.canvascmp.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # Infó sáv a két fájlhoz egymás mellett
        info_frame = tk.Frame(parent, bg="#eaf4fb")
        info_frame.pack(fill=tk.X)
        self._cmp_info1 = tk.StringVar(value="1. fájl: –")
        self._cmp_info2 = tk.StringVar(value="2. fájl: –")
        tk.Label(info_frame, textvariable=self._cmp_info1,
                 bg="#d6eaf8", fg="#1a5276", font=("Segoe UI", 9, "bold"),
                 anchor="w", padx=10, pady=3,
                 ).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(info_frame, textvariable=self._cmp_info2,
                 bg="#d5f5e3", fg="#196f3d", font=("Segoe UI", 9, "bold"),
                 anchor="w", padx=10, pady=3,
                 ).pack(side=tk.LEFT, fill=tk.X, expand=True)

    # ── Tab 6 – Δ0-hoz magassághatár vezérlő ───────────────────────────────────
    def _init_tab6_controls(self, parent):
        cbar = tk.Frame(parent, bg="#f7f7f7", pady=6, padx=10, relief=tk.GROOVE, bd=1)
        cbar.pack(fill=tk.X)

        tk.Label(cbar, text="Felső magassághatár:",
                 bg="#f7f7f7", fg="#333", font=("Segoe UI", 10),
                 ).pack(side=tk.LEFT)

        self._delta_limit_scale = tk.Scale(
            cbar, variable=self._delta_limit_var,
            from_=0.0, to=1.0, resolution=0.01,
            orient=tk.HORIZONTAL, length=520,
            command=self._on_delta_limit_change,
            bg="#f7f7f7", troughcolor="#d9d9d9",
            highlightthickness=0, showvalue=False,
        )
        self._delta_limit_scale.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8)
        self._delta_limit_scale.config(state=tk.DISABLED)

        tk.Label(cbar, textvariable=self._delta_limit_label,
                 bg="#f7f7f7", fg="#1a5276", font=("Segoe UI", 10, "bold"),
                 ).pack(side=tk.LEFT, padx=8)

    # ── Tab 9 – Magasság pontfelhő ───────────────────────────────────────────────
    def _init_tab9_height(self, parent):
        cbar = tk.Frame(parent, bg="#f7f7f7", pady=6, padx=10, relief=tk.GROOVE, bd=1)
        cbar.pack(fill=tk.X)

        left_side = tk.Frame(cbar, bg="#f7f7f7")
        left_side.pack(side=tk.LEFT, fill=tk.X, expand=True)

        right_side = tk.Frame(cbar, bg="#f7f7f7")
        right_side.pack(side=tk.RIGHT)

        tk.Button(left_side, text="📁  Hónap mappa kiválasztása",
                  command=self.select_height_folder,
                  bg="#8b0000", fg="white", font=("Segoe UI", 10, "bold"),
                  relief=tk.FLAT, padx=10, pady=6, cursor="hand2",
                  ).pack(side=tk.LEFT)

        tk.Label(left_side, textvariable=self._height_folder_var,
                 bg="#f7f7f7", fg="#333", font=("Segoe UI", 9),
                 ).pack(side=tk.LEFT, padx=12)

        self._height_filter_btn = tk.Checkbutton(
            right_side,
            text="ZDR átlag szűrő",
            variable=self._height_filter_enabled,
            command=self._on_height_filter_toggle,
            bg="#f7f7f7",
            fg="#222",
            selectcolor="#dff0d8",
            activebackground="#f7f7f7",
            font=("Segoe UI", 10),
            relief=tk.FLAT,
            padx=8,
            pady=3,
            cursor="hand2",
            indicatoron=False,
            bd=1,
        )
        self._height_filter_btn.pack(side=tk.LEFT, padx=(0, 10))

        self._height_stats_btn = tk.Checkbutton(
            right_side,
            text="Intervallum statisztika",
            variable=self._height_stats_enabled,
            command=self._on_height_stats_toggle,
            bg="#f7f7f7",
            fg="#222",
            selectcolor="#dff0d8",
            activebackground="#f7f7f7",
            font=("Segoe UI", 10),
            relief=tk.FLAT,
            padx=8,
            pady=3,
            cursor="hand2",
            indicatoron=False,
            bd=1,
        )
        self._height_stats_btn.pack(side=tk.LEFT, padx=(0, 10))

        tk.Label(right_side, text="Alsó vágás:", bg="#f7f7f7", fg="#333",
                 font=("Segoe UI", 9)).pack(side=tk.LEFT)
        self._height_filter_low_spin = tk.Spinbox(
            right_side, from_=-100.0, to=100.0, increment=0.5,
            textvariable=self._height_filter_low_var, width=8,
            format="%.1f", state=tk.DISABLED,
            command=self._refresh_height_cloud_from_controls,
        )
        self._height_filter_low_spin.pack(side=tk.LEFT, padx=(6, 12))
        self._height_filter_low_spin.bind("<Return>", self._refresh_height_cloud_from_controls)
        self._height_filter_low_spin.bind("<FocusOut>", self._refresh_height_cloud_from_controls)

        tk.Label(right_side, text="Felső vágás:", bg="#f7f7f7", fg="#333",
                 font=("Segoe UI", 9)).pack(side=tk.LEFT)
        self._height_filter_high_spin = tk.Spinbox(
            right_side, from_=-100.0, to=100.0, increment=0.5,
            textvariable=self._height_filter_high_var, width=8,
            format="%.1f", state=tk.DISABLED,
            command=self._refresh_height_cloud_from_controls,
        )
        self._height_filter_high_spin.pack(side=tk.LEFT, padx=(6, 8))
        self._height_filter_high_spin.bind("<Return>", self._refresh_height_cloud_from_controls)
        self._height_filter_high_spin.bind("<FocusOut>", self._refresh_height_cloud_from_controls)

        self._height_filter_hint = tk.StringVar(
            value="A szűrő alapértelmezésben kikapcsolva van."
        )
        tk.Label(right_side, textvariable=self._height_filter_hint,
                 bg="#f7f7f7", fg="#666", font=("Segoe UI", 9),
                 anchor="w").pack(side=tk.LEFT, padx=8)

        tk.Label(parent, textvariable=self._height_status_var,
                 bg="#f0f8ff", fg="#1a5276", font=("Segoe UI", 9),
                 anchor="w", padx=10, pady=5,
                 ).pack(fill=tk.X)
        tk.Label(parent, textvariable=self._height_stats_var,
                 bg="#f0f8ff", fg="#1a5276", font=("Segoe UI", 9),
                 anchor="w", padx=10, pady=2,
                 ).pack(fill=tk.X)

        # Intervallum-választó gombok (adatbetöltés után töltődik fel)
        self._height_interval_frame = tk.Frame(parent, bg="#f8f8f8", pady=4, padx=8,
                                               relief=tk.GROOVE, bd=1)
        self._height_interval_frame.pack(fill=tk.X)
        self._height_selected_interval = tk.IntVar(value=-1)

        # Két subplot figure
        fig = Figure(figsize=(12, 6), dpi=96, facecolor="white")
        canvas = FigureCanvasTkAgg(fig, master=parent)
        canvas.draw()
        tb = tk.Frame(parent, bg="white")
        tb.pack(fill=tk.X)
        NavigationToolbar2Tk(canvas, tb).update()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self.fig6, self.canvas6 = fig, canvas

    def _init_tab12_table_export(self, parent):
        cbar = tk.Frame(parent, bg="#f7f7f7", pady=8, padx=10,
                        relief=tk.GROOVE, bd=1)
        cbar.pack(fill=tk.X)

        tk.Button(cbar, text="📁  Kalibrálás előtti mappa",
              command=self.select_table_before_folder,
              bg="#8b0000", fg="white", font=("Segoe UI", 10, "bold"),
              relief=tk.FLAT, padx=10, pady=6, cursor="hand2").pack(side=tk.LEFT)
        tk.Label(cbar, textvariable=self._table_before_folder_var, bg="#f7f7f7",
             fg="#555", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(6, 18))

        tk.Button(cbar, text="📁  Kalibrálás utáni mappa",
              command=self.select_table_after_folder,
                  bg="#6c0f0f", fg="white", font=("Segoe UI", 10, "bold"),
                  relief=tk.FLAT, padx=10, pady=6, cursor="hand2").pack(side=tk.LEFT)
        tk.Label(cbar, textvariable=self._table_after_folder_var,
                 bg="#f7f7f7", fg="#555", font=("Segoe UI", 9)).pack(
                     side=tk.LEFT, padx=10)

        tk.Button(cbar, text="💾  Exportálás táblázatba",
                  command=self.export_zdr_table,
                  bg="#21618c", fg="white", font=("Segoe UI", 10, "bold"),
                  relief=tk.FLAT, padx=10, pady=6, cursor="hand2").pack(
                      side=tk.RIGHT)

        tk.Button(cbar, text="📄  XLSX sablon",
                  command=self.select_table_template,
                  bg="#555555", fg="white", font=("Segoe UI", 9, "bold"),
                  relief=tk.FLAT, padx=8, pady=6, cursor="hand2").pack(
                      side=tk.RIGHT, padx=(0, 8))

        tk.Label(parent, textvariable=self._table_export_status_var,
                 bg="#f0f8ff", fg="#1a5276", font=("Segoe UI", 9),
                 anchor="w", padx=10, pady=6).pack(fill=tk.X)

        tk.Label(parent,
                 text="A táblázat csak az átlagos ZDR értékeket tartalmazza, 0–4 km és 5–35 dBZ között.",
                 bg="white", fg="#555", font=("Segoe UI", 10),
                 anchor="w", padx=10, pady=10).pack(fill=tk.X)

    def select_table_before_folder(self):
        folder = filedialog.askdirectory(
            title="Válassza ki a kalibrálás előtti hónap mappáját"
        )
        if not folder:
            return
        self._table_before_folder = folder
        self._table_before_folder_var.set(folder)
        self._table_export_status_var.set("A kalibrálás előtti mappa betöltve.")

    def select_table_after_folder(self):
        folder = filedialog.askdirectory(
            title="Válassza ki a kalibrálás utáni hónap mappáját"
        )
        if not folder:
            return
        self._table_after_folder = folder
        self._table_after_folder_var.set(folder)
        self._table_export_status_var.set(
            "Mindkét mappa kiválasztva. Az exportáláshoz kattintson az Exportálás táblázatba gombra."
        )

    def select_table_template(self):
        path = filedialog.askopenfilename(
            title="Válassza ki az XLSX export sablont",
            filetypes=[("Excel munkafüzet", "*.xlsx"), ("Minden fájl", "*.*")],
        )
        if path:
            self._table_template_path = path
            self._table_export_status_var.set(f"XLSX sablon: {path}")

    def _collect_zdr_table_data(self, root_folder):
        heights = []
        zdr_values = []
        z_values = []
        dates = []
        station = None

        for day_name in sorted(os.listdir(root_folder)):
            day_path = os.path.join(root_folder, day_name)
            if not os.path.isdir(day_path):
                continue
            for fname in sorted(os.listdir(day_path)):
                if not fname.lower().endswith(".txt"):
                    continue
                file_path = os.path.join(day_path, fname)
                if not os.path.isfile(file_path):
                    continue

                date_value = self._parse_datetime_from_filename(fname)
                if date_value is not None:
                    dates.append(date_value)
                if station is None and len(fname) >= 3:
                    station = _STATIONS.get(fname[:3].upper())

                file_heights, file_zdr, file_z = parse_file(file_path)
                if file_heights.size == 0:
                    continue
                heights.extend(file_heights.tolist())
                zdr_values.extend(file_zdr.tolist())
                z_values.extend(file_z.tolist())

        if not heights:
            return np.array([]), np.array([]), np.array([]), dates, station
        return (np.asarray(heights, dtype=float),
                np.asarray(zdr_values, dtype=float),
                np.asarray(z_values, dtype=float), dates, station)

    def _build_zdr_table_block(self, root_folder, calibration_label):
        if not root_folder or not os.path.isdir(root_folder):
            raise ValueError(f"A {calibration_label} mappa nincs kiválasztva vagy nem létezik.")

        heights, zdr_values, z_values, dates, station = self._collect_zdr_table_data(root_folder)
        if heights.size == 0:
            raise ValueError(f"A {calibration_label} mappában nem található adat.")

        matrix, _ = self._calculate_zdr_table_stats(heights, zdr_values, z_values)
        height_bins = [(idx * 0.5, (idx + 1) * 0.5) for idx in range(8)]
        z_bins = [(idx, idx + 5) for idx in range(0, 35, 5)]
        rows = []
        for row_index, (z_low, z_high) in enumerate(z_bins):
            row = [f"{z_low}-{z_high}dBz"]
            row.extend("" if np.isnan(value) else f"{value:.4f}"
                       for value in matrix[row_index])
            rows.append(row + [""])

        if dates:
            timestamps = np.asarray([date_value.timestamp() for date_value in dates], dtype=float)
            middle_date = datetime.fromtimestamp(float(np.median(timestamps)))
            date_label = f"{middle_date.year}, {middle_date.month:02d}"
        else:
            date_label = ""

        station_label = station or "Radar"
        return [
            [f"{station_label}, {calibration_label}"] + [""] * 8 + [
                "Év, hónap"
            ],
            ["h", "0-500m", "500-1000m", "1000-1500m", "1500-2000m",
             "2000-2500m", "2500-3000m", "3000-3500m", "3500-4000m", date_label],
            ["Z"] + [""] * 8 + ["Olvadási réteg"],
            *rows,
        ]

    def _table_data_middle_date(self, dates):
        if not dates:
            return None
        timestamps = np.asarray([date_value.timestamp() for date_value in dates], dtype=float)
        return datetime.fromtimestamp(float(np.median(timestamps)))

    def _calculate_zdr_table_matrix(self, heights, zdr_values, z_values):
        matrix, _ = self._calculate_zdr_table_stats(heights, zdr_values, z_values)
        return matrix

    def _calculate_zdr_table_stats(self, heights, zdr_values, z_values):
        """Kétlépcsős ZDR-átlag és a cellák nyers ZDR-elemszáma."""
        matrix = np.full((7, 8), np.nan, dtype=float)
        counts = np.zeros((7, 8), dtype=int)
        for z_index, z_low in enumerate(range(0, 35, 5)):
            z_high = z_low + 5
            z_last = z_high == 35
            z_upper_mask = z_values <= z_high if z_last else z_values < z_high
            z_mask = ((z_values >= z_low) & z_upper_mask)
            for height_index in range(8):
                h_low = height_index * 0.5
                h_high = h_low + 0.5
                h_last = h_high == 4.0
                height_upper_mask = heights <= h_high if h_last else heights < h_high
                cell_mask = (z_mask & (heights >= h_low) & height_upper_mask &
                             ~np.isnan(zdr_values))
                if not cell_mask.any():
                    continue

                average, count = self._average_zdr_by_height(
                    heights, zdr_values, cell_mask
                )
                counts[z_index, height_index] = count
                if not np.isnan(average):
                    matrix[z_index, height_index] = average
        return matrix, counts

    def _average_zdr_by_height(self, heights, zdr_values, base_mask):
        """Egy maszkelt cella kétlépcsős ZDR-átlaga és nyers pontszáma."""
        valid_mask = base_mask & ~np.isnan(heights) & ~np.isnan(zdr_values)
        if not valid_mask.any():
            return np.nan, 0

        selected_heights = heights[valid_mask]
        selected_zdr = zdr_values[valid_mask]
        height_means = []
        for height in np.unique(selected_heights):
            values_at_height = selected_zdr[selected_heights == height]
            if values_at_height.size:
                height_means.append(float(np.mean(values_at_height)))
        if not height_means:
            return np.nan, 0
        return float(np.mean(height_means)), int(selected_zdr.size)

    def _build_zdr_difference_block(self, before_matrix, after_matrix, station):
        height_labels = ["0-500m", "500-1000m", "1000-1500m", "1500-2000m",
                         "2000-2500m", "2500-3000m", "3000-3500m", "3500-4000m"]
        rows = []
        for row_index, z_low in enumerate(range(0, 35, 5)):
            row = [f"{z_low}-{z_low + 5}dBz"]
            for before_value, after_value in zip(before_matrix[row_index], after_matrix[row_index]):
                if np.isnan(before_value) or np.isnan(after_value):
                    row.append("")
                else:
                    row.append(f"{abs(after_value) - abs(before_value):+.4f}")
            rows.append(row + [""])
        return [
            [f"{station or 'Radar'}, kalibrálás utáni - előtti: 0-hoz viszonyított változás"] + [""] * 9,
            ["h", *height_labels, ""],
            ["Z"] + [""] * 8 + ["+ távolodás", "- közeledés"],
            *rows,
        ]

    def export_zdr_table(self):
        if not _OPENPYXL_OK:
            messagebox.showerror(
                "Hiányzó csomag",
                "Az XLSX exporthoz telepítse az openpyxl csomagot: pip install openpyxl"
            )
            return

        pre_folder = self._table_before_folder
        post_folder = self._table_after_folder
        if not pre_folder:
            messagebox.showwarning(
                "Hiányzó mappa",
            "Előbb válassza ki a kalibrálás előtti mappát az export fülön."
            )
            return
        if not post_folder:
            messagebox.showwarning(
                "Hiányzó mappa",
                "Előbb válassza ki a kalibrálás utáni mappát az export fülön."
            )
            return

        try:
            before_data = self._collect_zdr_table_data(pre_folder)
            after_data = self._collect_zdr_table_data(post_folder)
            before_date = self._table_data_middle_date(before_data[3])
            after_date = self._table_data_middle_date(after_data[3])
            if before_date is not None and after_date is not None and after_date < before_date:
                before_folder, after_folder = post_folder, pre_folder
                before_data, after_data = after_data, before_data
                before_date, after_date = after_date, before_date
            else:
                before_folder, after_folder = pre_folder, post_folder

            pre_block = self._build_zdr_table_block(before_folder, "kalibrálás előtt")
            post_block = self._build_zdr_table_block(after_folder, "kalibrálás után")
            before_matrix, before_counts = self._calculate_zdr_table_stats(*before_data[:3])
            after_matrix, after_counts = self._calculate_zdr_table_stats(*after_data[:3])
            station = before_data[4] or after_data[4]
            difference_block = self._build_zdr_difference_block(
                before_matrix, after_matrix, station
            )
        except Exception as exc:
            messagebox.showerror("Exportálási hiba", f"Nem sikerült kiszámítani a táblázatot:\n{exc}")
            return

        template_path = self._table_template_path
        if not os.path.isfile(template_path):
            template_path = filedialog.askopenfilename(
                title="Válassza ki az XLSX export sablont",
                filetypes=[("Excel munkafüzet", "*.xlsx"), ("Minden fájl", "*.*")],
            )
        if not template_path:
            return

        output_path = filedialog.asksaveasfilename(
            title="ZDR átlagok exportálása Excel-fájlba",
            defaultextension=".xlsx",
            filetypes=[("Excel munkafüzet", "*.xlsx"), ("Minden fájl", "*.*")],
        )
        if not output_path:
            return

        try:
            workbook = load_workbook(template_path)
            worksheet = workbook.active
            self._write_table_block_to_sheet(worksheet, pre_block, 1, before_counts)
            self._write_table_block_to_sheet(worksheet, post_block, 11, after_counts)
            worksheet.cell(2, 10).value = pre_block[1][9]
            worksheet.cell(12, 10).value = post_block[1][9]
            self._write_difference_block_to_sheet(
                worksheet, difference_block, before_matrix, after_matrix, 21, 11
            )
            workbook.save(output_path)
        except Exception as exc:
            messagebox.showerror("Mentési hiba", f"Nem sikerült elmenteni a táblázatot:\n{exc}")
            return

        self._table_export_status_var.set(f"Táblázat exportálva: {output_path}")
        messagebox.showinfo("Exportálás", f"A színezett ZDR Excel-táblázat elkészült:\n{output_path}")

    def _write_table_block_to_sheet(self, worksheet, block, start_row, counts):
        for row_offset, values in enumerate(block):
            for column_index, value in enumerate(values[:9], start=1):
                cell = worksheet.cell(start_row + row_offset, column_index)
                if isinstance(cell, MergedCell):
                    continue
                cell.value = value

        strong_red_fill = PatternFill(fill_type="solid", fgColor="FCE4D6")
        light_red_fill = PatternFill(fill_type="solid", fgColor="FFF5F1")
        yellow_fill = PatternFill(fill_type="solid", fgColor="FFFEF2")
        lightest_green_fill = PatternFill(fill_type="solid", fgColor="F2F9EE")
        light_green_fill = PatternFill(fill_type="solid", fgColor="E2F0D9")
        for row_index in range(counts.shape[0]):
            for column_index in range(counts.shape[1]):
                cell = worksheet.cell(start_row + 3 + row_index, 2 + column_index)
                if isinstance(cell, MergedCell) or cell.value in (None, ""):
                    continue
                count = int(counts[row_index, column_index])
                if count < 10:
                    cell.fill = strong_red_fill
                elif count < 20:
                    cell.fill = light_red_fill
                elif count < 30:
                    cell.fill = yellow_fill
                elif count <= 50:
                    cell.fill = lightest_green_fill
                else:
                    cell.fill = light_green_fill

    def _write_difference_block_to_sheet(
            self, worksheet, block, before_matrix, after_matrix, start_row, style_source_row):
        for row_offset, values in enumerate(block):
            target_row = start_row + row_offset
            source_row = style_source_row + row_offset
            for column_index, value in enumerate(values[:9], start=1):
                target_cell = worksheet.cell(target_row, column_index)
                source_cell = worksheet.cell(source_row, column_index)
                if isinstance(target_cell, MergedCell):
                    continue
                original_fill = copy(target_cell.fill)
                target_cell._style = copy(source_cell._style)
                target_cell.fill = original_fill
                target_cell.value = value

        light_green_fill = PatternFill(fill_type="solid", fgColor="F2F9EE")
        green_fill = PatternFill(fill_type="solid", fgColor="E2F0D9")
        light_red_fill = PatternFill(fill_type="solid", fgColor="FFF5F1")
        red_fill = PatternFill(fill_type="solid", fgColor="FCE4D6")
        for row_index in range(before_matrix.shape[0]):
            for column_index in range(before_matrix.shape[1]):
                cell = worksheet.cell(start_row + 3 + row_index, 2 + column_index)
                before_value = before_matrix[row_index, column_index]
                after_value = after_matrix[row_index, column_index]
                if np.isnan(before_value) or np.isnan(after_value):
                    continue
                change = abs(after_value) - abs(before_value)
                if change < -0.1:
                    cell.fill = green_fill
                elif change < 0:
                    cell.fill = light_green_fill
                elif change > 0.1:
                    cell.fill = red_fill
                elif change > 0:
                    cell.fill = light_red_fill

    def _set_height_filter_widgets_state(self, state):
        for widget in (self._height_filter_low_spin, self._height_filter_high_spin):
            widget.config(state=state)

    def _on_height_filter_toggle(self):
        enabled = self._height_filter_enabled.get()
        self._set_height_filter_widgets_state(tk.NORMAL if enabled else tk.DISABLED)
        if hasattr(self, "_height_cloud_heights") and self._height_cloud_heights is not None:
            self._draw_height_cloud(self._height_selected_interval.get())
            self._attach_cursors()

    def _refresh_height_cloud_from_controls(self, *_args):
        if hasattr(self, "_height_cloud_heights") and self._height_cloud_heights is not None:
            self._draw_height_cloud(self._height_selected_interval.get())
            self._attach_cursors()

    def _height_filter_mask(self, zdr_values):
        if not self._height_filter_enabled.get():
            return np.ones_like(zdr_values, dtype=bool)
        try:
            low = float(self._height_filter_low_var.get())
            high = float(self._height_filter_high_var.get())
        except Exception:
            return np.ones_like(zdr_values, dtype=bool)
        if low > high:
            low, high = high, low
        return (~np.isnan(zdr_values)) & (zdr_values >= low) & (zdr_values <= high)

    def _on_height_stats_toggle(self):
        if self._height_stats_enabled.get():
            self._height_stats_var.set(
                "Mozgassa az egérkurzort a bal oldali ZDR grafikonon egy 500 méteres intervallum kijelöléséhez."
            )
            self._connect_height_stats_events()
        else:
            self._height_stats_var.set("Intervallum statisztika kikapcsolva.")
            self._height_stats_range = None
            self._disconnect_height_stats_events()
            if hasattr(self, '_height_cloud_heights'):
                self._draw_height_cloud(self._height_selected_interval.get())

    def _connect_height_stats_events(self):
        if getattr(self, '_height_stats_cid', None) is not None:
            return
        if hasattr(self, 'canvas6') and self.canvas6 is not None:
            self._height_stats_cid = self.canvas6.mpl_connect(
                'motion_notify_event', self._on_height_stats_move)

    def _disconnect_height_stats_events(self):
        if getattr(self, '_height_stats_cid', None) is None:
            return
        try:
            self.canvas6.mpl_disconnect(self._height_stats_cid)
        except Exception:
            pass
        self._height_stats_cid = None

    def _on_height_stats_move(self, event):
        if not self._height_stats_enabled.get() or event.inaxes is not self._height_cloud_ax1:
            return
        if event.xdata is None:
            return

        low = math.floor(event.xdata * 2.0) / 2.0
        high = low + 0.5
        if high <= low:
            high = low + 1

        if self._height_stats_range == (low, high):
            return
        self._height_stats_range = (low, high)
        self._compute_height_stats()
        self._draw_height_cloud(self._height_selected_interval.get())

    def _compute_height_stats(self):
        if self._height_stats_range is None:
            self._height_stats_var.set("Intervallum statisztika kikapcsolva.")
            return
        low, high = self._height_stats_range
        sel_idx = getattr(self, '_height_selected_interval', None)
        if sel_idx is None or sel_idx.get() < 0:
            self._height_stats_var.set(
                "Válasszon ki egy színes átlagolt intervallumot a statisztikához."
            )
            return

        if not hasattr(self, '_height_current_avg_maps') or not self._height_current_avg_maps:
            self._height_stats_var.set("Nincs elérhető aggregált ZDR adat.")
            return

        sel = sel_idx.get()
        if sel >= len(self._height_current_avg_maps) or sel < 0:
            self._height_stats_var.set("Érvénytelen intervallum kiválasztás.")
            return

        avg_map = self._height_current_avg_maps[sel]
        if not avg_map:
            self._height_stats_var.set(
                "Nincs elérhető átlagolt pont a kiválasztott intervallumban."
            )
            return

        interval_values = [value for h, value in avg_map.items() if low <= h < high]
        zdr_point_count = 0
        if hasattr(self, "_height_cloud_heights") and hasattr(self, "_height_cloud_zdr"):
            point_mask = (
                self._height_interval_masks[sel]
                & self._height_filter_mask(self._height_cloud_zdr)
                & (self._height_cloud_heights >= low)
                & (self._height_cloud_heights < high)
                & ~np.isnan(self._height_cloud_zdr)
            )
            zdr_point_count = int(np.count_nonzero(point_mask))
        if not interval_values:
            self._height_stats_var.set(
                f"Intervallum: {low}–{high} km  |  nincs átlagolt ZDR pont ebben a sávban "
                f"(ZDR pontok száma: N = {zdr_point_count})."
            )
            return

        mean_val = float(np.nanmean(interval_values))
        min_val = float(np.nanmin(interval_values))
        max_val = float(np.nanmax(interval_values))
        self._height_stats_var.set(
            f"Intervallum: {low}–{high} km   |   Átlag ZDR: {mean_val:.3f}   "
            f"Min: {min_val:.3f}   Max: {max_val:.3f}   |   "
            f"ZDR pontok száma: N = {zdr_point_count}"
        )

    # ── Tab 10 – Data Quality Control ───────────────────────────────────────
    def _init_tab10_dqc(self, parent):
        cbar = tk.Frame(parent, bg="#f7f7f7", pady=6, padx=10, relief=tk.GROOVE, bd=1)
        cbar.pack(fill=tk.X)

        left_side = tk.Frame(cbar, bg="#f7f7f7")
        left_side.pack(side=tk.LEFT, fill=tk.X, expand=True)

        right_side = tk.Frame(cbar, bg="#f7f7f7")
        right_side.pack(side=tk.RIGHT)

        tk.Button(left_side, text="📁  DQC mappa kiválasztása",
                  command=self.select_dqc_folder,
                  bg="#6c0f0f", fg="white", font=("Segoe UI", 10, "bold"),
                  relief=tk.FLAT, padx=10, pady=6, cursor="hand2",
                  ).pack(side=tk.LEFT)

        tk.Label(left_side, textvariable=self._dqc_folder_var,
                 bg="#f7f7f7", fg="#333", font=("Segoe UI", 9),
                 ).pack(side=tk.LEFT, padx=12)

        tk.Label(right_side, text="Magasság:", bg="#f7f7f7", fg="#333", font=("Segoe UI", 9)
                 ).pack(side=tk.LEFT)
        self._dqc_height_low_spin = tk.Spinbox(
            right_side, from_=-100.0, to=100000.0, increment=0.1,
            textvariable=self._dqc_height_low_var, width=7,
            format="%.2f", command=self._run_dqc_filter
        )
        self._dqc_height_low_spin.pack(side=tk.LEFT, padx=(6, 4))
        tk.Label(right_side, text="…", bg="#f7f7f7", fg="#333", font=("Segoe UI", 9)
                 ).pack(side=tk.LEFT)
        self._dqc_height_high_spin = tk.Spinbox(
            right_side, from_=-100.0, to=100000.0, increment=0.1,
            textvariable=self._dqc_height_high_var, width=7,
            format="%.2f", command=self._run_dqc_filter
        )
        self._dqc_height_high_spin.pack(side=tk.LEFT, padx=(4, 12))

        tk.Label(right_side, text="(ZH+ZV)/2:", bg="#f7f7f7", fg="#333", font=("Segoe UI", 9)
                 ).pack(side=tk.LEFT)
        self._dqc_zhzv_low_spin = tk.Spinbox(
            right_side, from_=-50.0, to=100.0, increment=0.5,
            textvariable=self._dqc_zhzv_low_var, width=7,
            format="%.2f", command=self._run_dqc_filter
        )
        self._dqc_zhzv_low_spin.pack(side=tk.LEFT, padx=(6, 4))
        tk.Label(right_side, text="…", bg="#f7f7f7", fg="#333", font=("Segoe UI", 9)
                 ).pack(side=tk.LEFT)
        self._dqc_zhzv_high_spin = tk.Spinbox(
            right_side, from_=-50.0, to=100.0, increment=0.5,
            textvariable=self._dqc_zhzv_high_var, width=7,
            format="%.2f", command=self._run_dqc_filter
        )
        self._dqc_zhzv_high_spin.pack(side=tk.LEFT, padx=(4, 12))

        tk.Label(right_side, text="RHOHV >=", bg="#f7f7f7", fg="#333", font=("Segoe UI", 9)
                 ).pack(side=tk.LEFT)
        self._dqc_rhohv_spin = tk.Spinbox(
            right_side, from_=0.0, to=1.0, increment=0.01,
            textvariable=self._dqc_rhohv_min_var, width=6,
            format="%.2f", command=self._run_dqc_filter
        )
        self._dqc_rhohv_spin.pack(side=tk.LEFT, padx=(6, 12))

        tk.Label(right_side, text="Átlag SNR >", bg="#f7f7f7", fg="#333", font=("Segoe UI", 9)
                 ).pack(side=tk.LEFT)
        self._dqc_snr_spin = tk.Spinbox(
            right_side, from_=-50.0, to=100.0, increment=0.5,
            textvariable=self._dqc_snr_min_var, width=7,
            format="%.2f", command=self._run_dqc_filter
        )
        self._dqc_snr_spin.pack(side=tk.LEFT, padx=(6, 8))

        tk.Button(right_side, text="Szűrés futtatása", command=self._run_dqc_filter,
                  bg="#21618c", fg="white", font=("Segoe UI", 9, "bold"),
                  relief=tk.FLAT, padx=10, pady=4, cursor="hand2"
                  ).pack(side=tk.LEFT)

        for widget in (self._dqc_height_low_spin, self._dqc_height_high_spin,
                   self._dqc_zhzv_low_spin, self._dqc_zhzv_high_spin,
                       self._dqc_rhohv_spin, self._dqc_snr_spin):
            widget.bind("<Return>", self._run_dqc_filter)
            widget.bind("<FocusOut>", self._run_dqc_filter)

        tk.Label(parent, textvariable=self._dqc_status_var,
                 bg="#f0f8ff", fg="#1a5276", font=("Segoe UI", 9),
                 anchor="w", padx=10, pady=5,
                 ).pack(fill=tk.X)

        list_wrap = tk.Frame(parent, bg="white")
        list_wrap.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        tk.Label(list_wrap, text="Szűrésen átment fájlok:",
                 bg="white", fg="#333", font=("Segoe UI", 10, "bold"), anchor="w"
                 ).pack(fill=tk.X, pady=(4, 4))

        lb_frame = tk.Frame(list_wrap, bg="white")
        lb_frame.pack(fill=tk.BOTH, expand=True)
        self._dqc_listbox = tk.Listbox(
            lb_frame,
            font=("Consolas", 9),
            bg="#fcfcfc",
            fg="#222",
            selectbackground="#d0eaff",
            selectforeground="#000",
            activestyle="none",
        )
        self._dqc_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb = tk.Scrollbar(lb_frame, orient=tk.VERTICAL, command=self._dqc_listbox.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self._dqc_listbox.configure(yscrollcommand=sb.set)

    def _init_tab11_melt_layer(self, parent):
        cbar = tk.Frame(parent, bg="#f7f7f7", pady=6, padx=10, relief=tk.GROOVE, bd=1)
        cbar.pack(fill=tk.X)

        left_side = tk.Frame(cbar, bg="#f7f7f7")
        left_side.pack(side=tk.LEFT, fill=tk.X, expand=True)

        right_side = tk.Frame(cbar, bg="#f7f7f7")
        right_side.pack(side=tk.RIGHT)

        tk.Button(left_side, text="📁  Év mappa kiválasztása",
                  command=self.select_melt_year_folder,
                  bg="#2f6f9f", fg="white", font=("Segoe UI", 10, "bold"),
                  relief=tk.FLAT, padx=10, pady=6, cursor="hand2",
                  ).pack(side=tk.LEFT)

        tk.Label(left_side, textvariable=self._melt_folder_var,
                 bg="#f7f7f7", fg="#333", font=("Segoe UI", 9),
                 ).pack(side=tk.LEFT, padx=12)

        self._melt_filter_btn = tk.Checkbutton(
            right_side,
            text="Érték szűrő",
            variable=self._melt_filter_enabled,
            command=self._on_melt_filter_toggle,
            bg="#f7f7f7",
            fg="#222",
            selectcolor="#dff0d8",
            activebackground="#f7f7f7",
            font=("Segoe UI", 10),
            relief=tk.FLAT,
            padx=8,
            pady=3,
            cursor="hand2",
            indicatoron=False,
            bd=1,
        )
        self._melt_filter_btn.pack(side=tk.LEFT, padx=(0, 10))

        tk.Label(right_side, text="Alsó vágás:", bg="#f7f7f7", fg="#333",
                 font=("Segoe UI", 9)).pack(side=tk.LEFT)
        self._melt_filter_low_spin = tk.Spinbox(
            right_side, from_=-1000.0, to=1000.0, increment=0.5,
            textvariable=self._melt_filter_low_var, width=8,
            format="%.1f", state=tk.DISABLED,
            command=self._refresh_melt_monthly_plot,
        )
        self._melt_filter_low_spin.pack(side=tk.LEFT, padx=(6, 12))
        self._melt_filter_low_spin.bind("<Return>", self._refresh_melt_monthly_plot)
        self._melt_filter_low_spin.bind("<FocusOut>", self._refresh_melt_monthly_plot)

        tk.Label(right_side, text="Felső vágás:", bg="#f7f7f7", fg="#333",
                 font=("Segoe UI", 9)).pack(side=tk.LEFT)
        self._melt_filter_high_spin = tk.Spinbox(
            right_side, from_=-1000.0, to=1000.0, increment=0.5,
            textvariable=self._melt_filter_high_var, width=8,
            format="%.1f", state=tk.DISABLED,
            command=self._refresh_melt_monthly_plot,
        )
        self._melt_filter_high_spin.pack(side=tk.LEFT, padx=(6, 8))
        self._melt_filter_high_spin.bind("<Return>", self._refresh_melt_monthly_plot)
        self._melt_filter_high_spin.bind("<FocusOut>", self._refresh_melt_monthly_plot)

        tk.Label(parent, textvariable=self._melt_status_var,
             bg="#f0f8ff", fg="#1a5276", font=("Segoe UI", 9),
             anchor="w", padx=10, pady=5,
             ).pack(fill=tk.X)

        tk.Button(right_side, text="💾 Mentés",
              command=self._save_melt_figure,
              bg="#21618c", fg="white", font=("Segoe UI", 9, "bold"),
              relief=tk.FLAT, padx=8, pady=4, cursor="hand2",
              ).pack(side=tk.LEFT, padx=(6, 8))

        fig = Figure(figsize=(10, 5), dpi=96, facecolor="white")
        self.fig7 = fig
        self.canvas7 = FigureCanvasTkAgg(fig, master=parent)
        self.canvas7.draw()
        self.canvas7.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        self._melt_year_plot_ax = self.fig7.add_subplot(111)

    def _extract_month_number(self, month_name):
        raw = str(month_name).strip().lower()
        raw = raw.replace("_", " ").replace("-", " ")
        raw = raw.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ö", "o").replace("ü", "u").replace("ű", "u")
        raw = raw.replace("ú", "u")

        month_map = {
            "januar": 1, "jan": 1,
            "februar": 2, "feb": 2,
            "marcius": 3, "mar": 3,
            "aprilis": 4, "apr": 4,
            "majus": 5, "may": 5,
            "junius": 6, "jun": 6,
            "julius": 7, "jul": 7,
            "augusztus": 8, "aug": 8,
            "szeptember": 9, "sep": 9,
            "oktober": 10, "oct": 10,
            "november": 11, "nov": 11,
            "december": 12, "dec": 12,
        }

        if not raw:
            return None

        for key, value in month_map.items():
            if raw == key:
                return value

        digits = re.findall(r"\d+", raw)
        if not digits:
            return None

        for token in digits:
            if len(token) == 1 and token.isdigit():
                month = int(token)
                if 1 <= month <= 12:
                    return month
            elif len(token) >= 2:
                month = int(token[:2]) if token[0] != '0' else int(token[1:3])
                if 1 <= month <= 12:
                    return month

        for key, value in month_map.items():
            if key in raw:
                return value

        return None

    def _parse_melt_value_from_text(self, text_line):
        if text_line is None:
            return None
        stripped = text_line.strip()
        if not stripped:
            return None

        match = re.search(r"^\s{4,}([-+]?\d+(?:[.,]\d+)?)", text_line)
        if match:
            return float(match.group(1).replace(",", "."))

        match = re.search(r"^\s+([-+]?\d+(?:[.,]\d+)?)", text_line)
        if match:
            return float(match.group(1).replace(",", "."))

        parts = stripped.split()
        if len(parts) >= 1:
            try:
                return float(parts[0].replace(",", "."))
            except Exception:
                pass
        return None

    def _read_melt_file_value(self, file_path):
        try:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    value = self._parse_melt_value_from_text(line)
                    if value is not None:
                        return float(value)
            return None
        except Exception:
            return None

    def _melt_filter_mask(self, value):
        if not self._melt_filter_enabled.get():
            return True
        try:
            low = float(self._melt_filter_low_var.get())
            high = float(self._melt_filter_high_var.get())
        except Exception:
            return True
        if low > high:
            low, high = high, low
        return low <= value <= high

    def _collect_melt_layer_monthly_values(self, root_folder):
        if not os.path.isdir(root_folder):
            raise FileNotFoundError(f"A megadott útvonal nem létezik: {root_folder}")

        month_values = {m: [] for m in range(1, 13)}
        month_seen = set()

        for item in sorted(os.listdir(root_folder), key=lambda s: s.lower()):
            item_path = os.path.join(root_folder, item)

            if os.path.isdir(item_path):
                month_num = self._extract_month_number(item)
                if month_num is None:
                    continue
                month_seen.add(month_num)

                files = [
                    os.path.join(item_path, f)
                    for f in sorted(os.listdir(item_path), key=lambda s: s.lower())
                    if f.lower().endswith(".txt") and os.path.isfile(os.path.join(item_path, f))
                ]
                for file_path in files:
                    value = self._read_melt_file_value(file_path)
                    if value is None:
                        continue
                    if not self._melt_filter_mask(value):
                        continue  # kizárva az átlagolásból és az elemszámból is
                    month_values[month_num].append(float(value))

                for sub in sorted(os.listdir(item_path), key=lambda s: s.lower()):
                    sub_path = os.path.join(item_path, sub)
                    if not os.path.isdir(sub_path):
                        continue
                    for file_name in sorted(os.listdir(sub_path), key=lambda s: s.lower()):
                        if not file_name.lower().endswith(".txt"):
                            continue
                        file_path = os.path.join(sub_path, file_name)
                        if not os.path.isfile(file_path):
                            continue
                        value = self._read_melt_file_value(file_path)
                        if value is None:
                            continue
                        if not self._melt_filter_mask(value):
                            continue  # kizárva az átlagolásból és az elemszámból is
                        month_values[month_num].append(float(value))

        if not month_seen:
            for file_name in sorted(os.listdir(root_folder), key=lambda s: s.lower()):
                file_path = os.path.join(root_folder, file_name)
                if not os.path.isfile(file_path) or not file_name.lower().endswith(".txt"):
                    continue
                month_num = self._extract_month_number(file_name)
                if month_num is None:
                    continue
                month_seen.add(month_num)
                value = self._read_melt_file_value(file_path)
                if value is not None and self._melt_filter_mask(value):
                    month_values[month_num].append(float(value))  # csak a megengedett értékek számítanak bele az átlagba

        monthly_means = []
        for m in range(1, 13):
            vals = month_values[m]
            if vals:
                monthly_means.append(float(np.mean(vals)))
            else:
                monthly_means.append(np.nan)

        if not month_seen:
            raise ValueError("A kiválasztott év mappában nem találtam hónap mappákat vagy közvetlen .txt fájlokat (pl. 01, 02, ... vagy Január, Február, ...).")

        return monthly_means, sorted(month_seen)

    def _set_melt_filter_widgets_state(self, state):
        for widget in (self._melt_filter_low_spin, self._melt_filter_high_spin):
            widget.config(state=state)

    def _on_melt_filter_toggle(self):
        enabled = self._melt_filter_enabled.get()
        self._set_melt_filter_widgets_state(tk.NORMAL if enabled else tk.DISABLED)
        if self._melt_folder is not None:
            self._refresh_melt_monthly_plot()

    def _refresh_melt_monthly_plot(self, *_args):
        if getattr(self, '_melt_folder', None) is None:
            return
        try:
            monthly_values, _ = self._collect_melt_layer_monthly_values(self._melt_folder)
            self._plot_melt_monthly_averages(monthly_values)
            self._melt_status_var.set("Éves havi összesítés elkészült.")
        except Exception as exc:
            self._melt_status_var.set(f"Hiba a szűrés során: {exc}")

    def _plot_melt_monthly_averages(self, monthly_values):
        self._melt_monthly_values = np.asarray(monthly_values, dtype=float)
        ax = self._melt_year_plot_ax
        ax.clear()
        months = ["Január", "Február", "Március", "Április", "Május", "Június",
                  "Július", "Augusztus", "Szeptember", "Október", "November", "December"]

        values = self._melt_monthly_values
        colors = ["#4c78a8" if not np.isnan(v) else "#d9d9d9" for v in values]
        ax.bar(months, values, color=colors, edgecolor="#333", linewidth=0.8)
        ax.set_title("Olvadási réteg magasságok havi átlagai", fontsize=11, fontweight="bold")
        ax.set_ylabel("Átlagos magasság", fontsize=10)
        ax.set_xlabel("Hónap", fontsize=10)
        ax.grid(True, axis="y", linestyle="--", alpha=0.4)
        ax.tick_params(axis='x', rotation=30)

        max_val = np.nanmax(values) if np.any(~np.isnan(values)) else 1.0
        for idx, val in enumerate(values):
            if np.isnan(val):
                continue
            ax.text(idx, val + 0.02 * max_val, f"{val:.2f}", ha="center", va="bottom", fontsize=8)

        self._melt_summary_var.set("")
        self.canvas7.draw()

    def select_melt_year_folder(self):
        folder = filedialog.askdirectory(
            title="Válasszon év mappát az olvadási réteg adatainak beolvasásához"
        )
        if not folder:
            return

        self._melt_folder = folder
        self._melt_folder_var.set(folder)
        self._melt_status_var.set("Éves mappa feldolgozása...")

        try:
            monthly_values, _ = self._collect_melt_layer_monthly_values(folder)
            self._plot_melt_monthly_averages(monthly_values)
            self._melt_status_var.set("Éves havi összesítés elkészült.")
        except Exception as exc:
            messagebox.showerror("Hiba", f"Nem sikerült beolvasni az éves mappastruktúrát:\n{exc}")
            self._melt_status_var.set("Hiba történt a beolvasáskor.")

    def _save_melt_figure(self):
        if not hasattr(self, 'fig7') or self.fig7 is None:
            messagebox.showwarning("Mentés", "Nincs menthető diagram.")
            return
        try:
            fpath = filedialog.asksaveasfilename(
                title="Mentés képként",
                defaultextension=".png",
                filetypes=[("PNG képfájl", "*.png"), ("PDF", "*.pdf"), ("SVG", "*.svg"), ("JPEG", "*.jpg")]
            )
            if not fpath:
                return
            self.fig7.savefig(fpath, dpi=300, bbox_inches='tight')
            messagebox.showinfo("Mentés", f"Kép sikeresen elmentve: {fpath}")
        except Exception as exc:
            messagebox.showerror("Mentés hiba", f"Nem sikerült elmenteni a képet:\n{exc}")

    # ── Tab 4 – Radarkép ──────────────────────────────────────────────────────
    def _init_tab4_radar(self, parent):
        top = tk.Frame(parent, bg="#16213e", pady=5, padx=8)
        top.pack(fill=tk.X)
        tk.Label(top, text="URL:", bg="#16213e", fg="#aaa",
                 font=("Segoe UI", 9)).pack(side=tk.LEFT)
        self._radar_url_var = tk.StringVar(value="–")
        tk.Label(top, textvariable=self._radar_url_var,
                 bg="#16213e", fg="#56b4e9", font=("Consolas", 9),
                 cursor="hand2").pack(side=tk.LEFT, padx=8)
        tk.Button(top, text="↻  Újratöltés", command=self._load_radar_image,
                  bg="#0078d4", fg="white", font=("Segoe UI", 9), relief=tk.FLAT,
                  padx=8, pady=3, cursor="hand2").pack(side=tk.RIGHT, padx=4)

        self._radar_canvas = tk.Canvas(parent, bg="#1a1a2e", highlightthickness=0)
        self._radar_canvas.pack(fill=tk.BOTH, expand=True)

        self._radar_status_var = tk.StringVar(
            value="Töltsön be egy kompatibilis .txt fájlt a radarkép megjelenítéséhez.")
        tk.Label(parent, textvariable=self._radar_status_var,
                 bg="#16213e", fg="#aaa", font=("Segoe UI", 9),
                 anchor="w", padx=8, pady=3).pack(fill=tk.X)

    # ── Radarkép betöltés ─────────────────────────────────────────────────────
    def _load_radar_image(self):
        if not self._radar_url:
            self._radar_status_var.set("Nincs érvényes URL – töltsön be kompatibilis fájlt.")
            return
        if not _PIL_OK:
            self._radar_status_var.set("Hiba: a Pillow csomag hiányzik.")
            return
        self._radar_status_var.set("Letöltés folyamatban…")
        self._radar_canvas.delete("all")
        threading.Thread(target=self._fetch_radar_image, daemon=True).start()

    def _fetch_radar_image(self):
        try:
            req = urllib.request.Request(
                self._radar_url,
                headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"},
            )
            with urllib.request.urlopen(req, timeout=15) as resp:
                data = resp.read()
            img = Image.open(io.BytesIO(data)).convert("RGBA")
            self.after(0, lambda: self._display_radar_image(img))
        except Exception as exc:
            msg = f"Nem sikerült letölteni a képet: {exc}"
            self.after(0, lambda: self._radar_status_var.set(msg))

    def _display_radar_image(self, img):
        c = self._radar_canvas
        c.update_idletasks()
        cw, ch = c.winfo_width(), c.winfo_height()
        if cw < 2 or ch < 2:
            cw, ch = 900, 560
        img_copy = img.copy()
        img_copy.thumbnail((cw, ch), Image.LANCZOS)
        photo = ImageTk.PhotoImage(img_copy)
        self._radar_image_ref = photo
        c.delete("all")
        c.create_image(cw // 2, ch // 2, anchor="center", image=photo)
        self._radar_status_var.set(f"Forrás: {self._radar_url}")

    # ── Tab 7 – Radarkép (2. állomás) ─────────────────────────────────────────
    def _init_tab7_radar(self, parent):
        top = tk.Frame(parent, bg="#16213e", pady=5, padx=8)
        top.pack(fill=tk.X)
        tk.Label(top, text="URL:", bg="#16213e", fg="#aaa",
                 font=("Segoe UI", 9)).pack(side=tk.LEFT)
        self._radar_url_var2 = tk.StringVar(value="–")
        tk.Label(top, textvariable=self._radar_url_var2,
                 bg="#16213e", fg="#56b4e9", font=("Consolas", 9),
                 cursor="hand2").pack(side=tk.LEFT, padx=8)
        tk.Button(top, text="↻  Újratöltés", command=self._load_radar_image2,
                  bg="#0078d4", fg="white", font=("Segoe UI", 9), relief=tk.FLAT,
                  padx=8, pady=3, cursor="hand2").pack(side=tk.RIGHT, padx=4)

        self._radar_canvas2 = tk.Canvas(parent, bg="#1a1a2e", highlightthickness=0)
        self._radar_canvas2.pack(fill=tk.BOTH, expand=True)

        self._radar_status_var2 = tk.StringVar(
            value="Töltsön be egy kompatibilis .txt fájlt (2. állomás) a radarkép megjelenítéséhez.")
        tk.Label(parent, textvariable=self._radar_status_var2,
                 bg="#16213e", fg="#aaa", font=("Segoe UI", 9),
                 anchor="w", padx=8, pady=3).pack(fill=tk.X)

    def _load_radar_image2(self):
        if not self._radar_url2:
            self._radar_status_var2.set("Nincs érvényes URL – töltsön be kompatibilis fájlt.")
            return
        if not _PIL_OK:
            self._radar_status_var2.set("Hiba: a Pillow csomag hiányzik.")
            return
        self._radar_status_var2.set("Letöltés folyamatban…")
        self._radar_canvas2.delete("all")
        threading.Thread(target=self._fetch_radar_image2, daemon=True).start()

    def _fetch_radar_image2(self):
        try:
            req = urllib.request.Request(
                self._radar_url2,
                headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"},
            )
            with urllib.request.urlopen(req, timeout=15) as resp:
                data = resp.read()
            img = Image.open(io.BytesIO(data)).convert("RGBA")
            self.after(0, lambda: self._display_radar_image2(img))
        except Exception as exc:
            msg = f"Nem sikerült letölteni a képet: {exc}"
            self.after(0, lambda: self._radar_status_var2.set(msg))

    def _display_radar_image2(self, img):
        c = self._radar_canvas2
        c.update_idletasks()
        cw, ch = c.winfo_width(), c.winfo_height()
        if cw < 2 or ch < 2:
            cw, ch = 900, 560
        img_copy = img.copy()
        img_copy.thumbnail((cw, ch), Image.LANCZOS)
        photo = ImageTk.PhotoImage(img_copy)
        self._radar_image_ref2 = photo
        c.delete("all")
        c.create_image(cw // 2, ch // 2, anchor="center", image=photo)
        self._radar_status_var2.set(f"Forrás: {self._radar_url2}")

    # ── Crosshair – Tab 1 & 2 ─────────────────────────────────────────────────
    def _attach_cursors(self):
        self._cursors.clear()
        for fig, canvas, prefix, x_is_date, hover_lookup in [
            (self.fig2, self.canvas2, "ZDR&Zátlag-grafikon", False, None),
            (self.fig3, self.canvas3, "Δ-különbségek", False, None),
            (getattr(self, 'fig4', None), getattr(self, 'canvas4', None), "Δ0-időbeli", False, None),
            (getattr(self, 'fig5', None), getattr(self, 'canvas5', None), "Havi ZDR pontfelhő", True, None),
            (getattr(self, 'fig6', None), getattr(self, 'canvas6', None), "Magasság pontfelhő", False, self._height_cloud_hover),
        ]:
            if fig and fig.axes:
                cur = CrosshairCursor(fig.axes[0], canvas, self._coord_var,
                                      label_prefix=prefix, x_is_date=x_is_date,
                                      hover_lookup=hover_lookup)
                cur.toggle(self._crosshair_on.get())
                self._cursors.append(cur)

    def _toggle_crosshair(self):
        state = self._crosshair_on.get()
        self._cursor_btn.config(bg="#0078d4" if state else "#f0f0f0",
                                fg="white" if state else "#222")
        if not state:
            self._coord_var.set("")
        for cur in self._cursors:
            cur.toggle(state)

    # ── Crosshair – Tab 3 ─────────────────────────────────────────────────────
    def _toggle_cmp_cursor(self):
        state = self._cmp_cursor_on.get()
        self._cmp_cursor_btn.config(bg="#217a3c" if state else "#f0f7f0",
                                    fg="white" if state else "#222")
        if not state:
            self._cmp_coord_var.set("")
        if self._cmp_cursor:
            self._cmp_cursor.toggle(state)

    def _init_tab8_month(self, parent):
        cbar = tk.Frame(parent, bg="#f7f7f7", pady=6, padx=10, relief=tk.GROOVE, bd=1)
        cbar.pack(fill=tk.X)

        left_side = tk.Frame(cbar, bg="#f7f7f7")
        left_side.pack(side=tk.LEFT, fill=tk.X, expand=True)

        right_side = tk.Frame(cbar, bg="#f7f7f7")
        right_side.pack(side=tk.RIGHT)

        tk.Button(left_side, text="📁  Hónap mappa kiválasztása",
                  command=self.select_month_folder,
                  bg="#8b0000", fg="white", font=("Segoe UI", 10, "bold"),
                  relief=tk.FLAT, padx=10, pady=6, cursor="hand2",
                  ).pack(side=tk.LEFT)

        tk.Label(left_side, textvariable=self._month_folder_var,
                 bg="#f7f7f7", fg="#333", font=("Segoe UI", 9),
                 ).pack(side=tk.LEFT, padx=12)

        self._month_daily_avg_btn = tk.Checkbutton(
            right_side,
            text="Napi átlagok",
            variable=self._month_daily_avg_enabled,
            command=self._on_month_daily_avg_toggle,
            bg="#f7f7f7",
            fg="#222",
            selectcolor="#dff0d8",
            activebackground="#f7f7f7",
            font=("Segoe UI", 10),
            relief=tk.FLAT,
            padx=8,
            pady=3,
            cursor="hand2",
            indicatoron=False,
            bd=1,
        )
        self._month_daily_avg_btn.pack(side=tk.LEFT, padx=(0, 8))

        self._month_break_btn = tk.Checkbutton(
            right_side,
            text="Törés",
            variable=self._month_break_enabled,
            command=self._on_month_break_toggle,
            bg="#f7f7f7",
            fg="#222",
            selectcolor="#dff0d8",
            activebackground="#f7f7f7",
            font=("Segoe UI", 10),
            relief=tk.FLAT,
            padx=8,
            pady=3,
            cursor="hand2",
            indicatoron=False,
            bd=1,
        )
        self._month_break_btn.pack(side=tk.LEFT)

        self._month_break_mode_btn = tk.Checkbutton(
            right_side,
            text="Töréspontos átlag",
            variable=self._month_break_mean_mode,
            command=self._on_month_break_toggle,
            bg="#f7f7f7",
            fg="#222",
            selectcolor="#dff0d8",
            activebackground="#f7f7f7",
            font=("Segoe UI", 10),
            relief=tk.FLAT,
            padx=8,
            pady=3,
            cursor="hand2",
            indicatoron=False,
            bd=1,
        )
        self._month_break_mode_btn.pack(side=tk.LEFT, padx=(8, 0))

        tk.Label(right_side, text="db:", bg="#f7f7f7", fg="#333", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(8, 4))
        self._month_break_count_spin = tk.Spinbox(
            right_side,
            from_=1,
            to=10,
            textvariable=self._month_break_count_var,
            width=3,
            command=self._on_month_break_toggle,
            state=tk.NORMAL,
            justify=tk.CENTER,
        )
        self._month_break_count_spin.pack(side=tk.LEFT)

        self._month_ci_btn = tk.Checkbutton(
            right_side,
            text="95% Konfidencia",
            variable=self._month_ci_enabled,
            command=self._on_month_ci_toggle,
            bg="#f7f7f7",
            fg="#222",
            selectcolor="#dff0d8",
            activebackground="#f7f7f7",
            font=("Segoe UI", 10),
            relief=tk.FLAT,
            padx=8,
            pady=3,
            cursor="hand2",
            indicatoron=False,
            bd=1,
        )
        self._month_ci_btn.pack(side=tk.LEFT, padx=(8, 0))

        tk.Label(parent, textvariable=self._month_status_var,
                 bg="#f0f8ff", fg="#1a5276", font=("Segoe UI", 9),
                 anchor="w", padx=10, pady=5,
                 ).pack(fill=tk.X)

    def _on_month_daily_avg_toggle(self):
        if hasattr(self, "_month_folder") and self._month_folder:
            self._plot_month_folder_tab(
                getattr(self, "_month_dates", []),
                getattr(self, "_month_zdr_values", np.array([])),
                getattr(self, "_month_zdr_samples", []),
            )
            self._attach_cursors()

    def _on_month_break_toggle(self):
        if not self._month_break_enabled.get():
            self._month_break_points = []
            self._month_break_x = None
            self._month_break_drag_index = None
            self._month_break_dragging = False
        if hasattr(self, "_month_folder") and self._month_folder:
            self._plot_month_folder_tab(
                getattr(self, "_month_dates", []),
                getattr(self, "_month_zdr_values", np.array([])),
                getattr(self, "_month_zdr_samples", []),
            )
            self._attach_cursors()

    def _disconnect_month_break_events(self):
        if not hasattr(self, "canvas5"):
            return
        for attr in ("_month_break_cid_press", "_month_break_cid_move", "_month_break_cid_release"):
            cid = getattr(self, attr, None)
            if cid is not None:
                try:
                    self.canvas5.mpl_disconnect(cid)
                except Exception:
                    pass
                setattr(self, attr, None)

    def _bind_month_break_events(self):
        self._disconnect_month_break_events()
        if not self._month_break_enabled.get():
            return
        self._month_break_cid_press = self.canvas5.mpl_connect("button_press_event", self._on_month_break_press)
        self._month_break_cid_move = self.canvas5.mpl_connect("motion_notify_event", self._on_month_break_move)
        self._month_break_cid_release = self.canvas5.mpl_connect("button_release_event", self._on_month_break_release)

    def _on_month_break_press(self, event):
        if not self._month_break_enabled.get() or event.inaxes is not self._month_break_ax or event.xdata is None:
            return
        if not self._month_break_points:
            return
        points = sorted(float(v) for v in self._month_break_points)
        xmin, xmax = self._month_break_xlim if self._month_break_xlim else self._month_break_ax.get_xlim()
        tol = max((xmax - xmin) * 0.03, 0.05)
        for idx, value in enumerate(points):
            if abs(event.xdata - value) <= tol:
                self._month_break_drag_index = idx
                self._month_break_dragging = True
                self._month_break_x = value
                return

    def _on_month_break_move(self, event):
        if (not self._month_break_enabled.get() or not self._month_break_dragging
                or event.inaxes is not self._month_break_ax or event.xdata is None):
            return
        if self._month_cluster_x.size < 4 or self._month_break_drag_index is None:
            return

        points = sorted(float(v) for v in self._month_break_points)
        if not points:
            return

        idx = self._month_break_drag_index
        min_x = float(np.min(self._month_cluster_x))
        max_x = float(np.max(self._month_cluster_x))
        span = max(max_x - min_x, 1.0)
        eps = max(span * 1e-6, 1e-9)

        lower_bound = min_x + eps if idx == 0 else points[idx - 1] + eps
        upper_bound = max_x - eps if idx == len(points) - 1 else points[idx + 1] - eps
        if upper_bound <= lower_bound:
            return

        new_x = min(max(float(event.xdata), lower_bound), upper_bound)
        points[idx] = new_x
        points = sorted(np.unique(np.round(points, 10)))
        self._month_break_points = [float(v) for v in points]
        self._month_break_x = new_x
        self._plot_month_folder_tab(
            self._month_dates, self._month_zdr_values, self._month_zdr_samples)

    def _on_month_break_release(self, _event):
        if not self._month_break_enabled.get():
            return
        was_dragging = self._month_break_dragging
        self._month_break_dragging = False
        self._month_break_drag_index = None
        if was_dragging:
            self._attach_cursors()

    def _parse_datetime_from_filename(self, filename):
        base = os.path.splitext(filename)[0]
        if len(base) < 15:
            return None
        digits = base[3:15]
        if not digits.isdigit():
            return None
        try:
            return datetime(
                int(digits[0:4]), int(digits[4:6]), int(digits[6:8]),
                int(digits[8:10]), int(digits[10:12])
            )
        except Exception:
            return None

    def select_month_folder(self):
        folder = filedialog.askdirectory(
            title="Válasszon hónap mappát a ZDR pontfelhőhöz"
        )
        if not folder:
            return
        self._month_folder = folder
        self._month_folder_var.set(folder)
        self._month_status_var.set("Hónap mappa betöltése...")

        try:
            dates, zdr_values, zdr_samples = self._collect_month_zdr(folder)
        except Exception as exc:
            messagebox.showerror("Hiba", f"Nem sikerült beolvasni a hónap mappát:\n{exc}")
            self._month_status_var.set("Hiba történt a hónap mappa beolvasásakor.")
            return

        if len(dates) == 0:
            self._month_status_var.set("Nem található érvényes ZDR adat a kiválasztott mappában.")
            self.fig5.clear()
            self.canvas5.draw()
            return

        self._plot_month_folder_tab(dates, zdr_values, zdr_samples)
        self._month_status_var.set(f"{len(dates)} pont betöltve.")

        self._height_folder = folder
        self._height_folder_var.set(folder)
        self._height_status_var.set("Hónap mappa betöltése...")
        try:
            heights, zdr_values, za_values = self._collect_height_cloud_data(folder)
            self._plot_height_folder_tab(heights, zdr_values, za_values)
            self._height_status_var.set(f"{heights.size} pont betöltve.")
        except Exception as exc:
            self._height_status_var.set("Nem sikerült a magasságpontfelhő betöltése.")

        self._attach_cursors()

    def select_height_folder(self):
        folder = filedialog.askdirectory(
            title="Válasszon hónap mappát a magasság alapú pontfelhőhöz"
        )
        if not folder:
            return
        self._height_folder = folder
        self._height_folder_var.set(folder)
        self._height_status_var.set("Hónap mappa betöltése...")

        try:
            heights, zdr_values, za_values = self._collect_height_cloud_data(folder)
        except Exception as exc:
            messagebox.showerror("Hiba", f"Nem sikerült beolvasni a hónap mappát:\n{exc}")
            self._height_status_var.set("Hiba történt a magasságpontfelhő beolvasásakor.")
            return

        if heights.size == 0:
            self._height_status_var.set(
                "Nem található érvényes magasságpont a kiválasztott mappában.")
            self.fig6.clear()
            self.canvas6.draw()
            return

        self._plot_height_folder_tab(heights, zdr_values, za_values)
        self._height_status_var.set(f"{heights.size} pont betöltve.")
        self._attach_cursors()

    def select_dqc_folder(self):
        folder = filedialog.askdirectory(
            title="Válasszon Data Quality Control mappát"
        )
        if not folder:
            return

        try:
            dqc_files = self._collect_dqc_files(folder)
        except Exception as exc:
            messagebox.showerror("Hiba", f"Nem sikerült beolvasni a DQC mappát:\n{exc}")
            self._dqc_status_var.set("A kiválasztott mappa nem megfelelő DQC szerkezetű.")
            return

        if not dqc_files:
            messagebox.showwarning(
                "Érvénytelen mappa",
                "A kiválasztott mappában nem található megfelelő Radar_neve -> *.txt szerkezet 12 oszlopos adatokkal."
            )
            self._dqc_status_var.set("Nincs megfelelő 12 oszlopos DQC adat a mappában.")
            return

        self._dqc_folder = folder
        self._dqc_files = dqc_files
        self._dqc_folder_var.set(folder)
        self._run_dqc_filter()

    def _collect_dqc_files(self, root_folder):
        if not os.path.isdir(root_folder):
            return []

        collected = []
        found_any_txt = False

        # 1) Gyökér -> Radar_mappa -> *.txt
        radar_dirs = [
            name for name in sorted(os.listdir(root_folder))
            if os.path.isdir(os.path.join(root_folder, name))
        ]
        if radar_dirs:
            for radar_name in radar_dirs:
                radar_path = os.path.join(root_folder, radar_name)
                for fname in sorted(os.listdir(radar_path)):
                    if not fname.lower().endswith(".txt"):
                        continue
                    found_any_txt = True
                    file_path = os.path.join(radar_path, fname)
                    if not os.path.isfile(file_path):
                        continue
                    if self._has_12col_numeric_rows(file_path):
                        collected.append((radar_name, fname, file_path))
                    else:
                        return []

        # 2) Közvetlen radar mappa -> *.txt (pl. BUD -> BUD2025...txt)
        if not found_any_txt:
            radar_name = os.path.basename(os.path.normpath(root_folder))
            for fname in sorted(os.listdir(root_folder)):
                if not fname.lower().endswith(".txt"):
                    continue
                found_any_txt = True
                file_path = os.path.join(root_folder, fname)
                if not os.path.isfile(file_path):
                    continue
                if self._has_12col_numeric_rows(file_path):
                    collected.append((radar_name, fname, file_path))
                else:
                    return []

        if not found_any_txt:
            return []
        return collected

    def _has_12col_numeric_rows(self, file_path):
        with open(file_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                parts = line.split()
                if len(parts) != 12:
                    continue
                try:
                    float(parts[0])
                except ValueError:
                    continue
                return True
        return False

    def _format_dqc_height_summary(self, passed_heights):
        if not passed_heights:
            return "magasság: n/a"

        vals = np.array(passed_heights, dtype=float)
        vals = vals[~np.isnan(vals)]
        if vals.size == 0:
            return "magasság: n/a"

        uniq = np.unique(np.round(vals, 3))
        if uniq.size == 1:
            return f"magasság: {uniq[0]:.3f} km"

        # Egymáshoz közeli pontokat egy tartományba vonunk.
        gap_km = 0.20
        ranges = []
        start = uniq[0]
        prev = uniq[0]
        for h in uniq[1:]:
            if (h - prev) > gap_km:
                ranges.append((start, prev))
                start = h
            prev = h
        ranges.append((start, prev))

        if len(ranges) == 1:
            lo, hi = ranges[0]
            return f"magasság tartomány: {lo:.3f}-{hi:.3f} km"

        segments = []
        for lo, hi in ranges[:3]:
            if np.isclose(lo, hi):
                segments.append(f"{lo:.3f}")
            else:
                segments.append(f"{lo:.3f}-{hi:.3f}")
        suffix = "" if len(ranges) <= 3 else " ..."
        return f"magasság tartományok: {', '.join(segments)} km{suffix}"

    def _dqc_file_passes(self, file_path, height_low, height_high, zhzv_low, zhzv_high, rhohv_min, snr_min):
        passed_rows = 0
        snr_pass_values = []
        passed_heights = []
        with open(file_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                parts = line.split()
                if len(parts) != 12:
                    continue
                try:
                    height = np.nan if parts[0].upper() == "NAN" else float(parts[0])
                    zh = np.nan if parts[2].upper() == "NAN" else float(parts[2])
                    zv = np.nan if parts[3].upper() == "NAN" else float(parts[3])
                    rhohv = np.nan if parts[7].upper() == "NAN" else float(parts[7])
                    snrv = np.nan if parts[8].upper() == "NAN" else float(parts[8])
                    snrh = np.nan if parts[9].upper() == "NAN" else float(parts[9])
                except ValueError:
                    continue

                if np.isnan(height) or height < height_low or height > height_high:
                    continue

                if np.isnan(zh) or np.isnan(zv):
                    continue
                dbz_avg = (zh + zv) / 2.0
                if dbz_avg < zhzv_low or dbz_avg > zhzv_high:
                    continue

                if np.isnan(rhohv) or rhohv < rhohv_min:
                    continue

                if np.isnan(snrh) or np.isnan(snrv):
                    continue
                row_snr_avg = (snrh + snrv) / 2.0
                if row_snr_avg > snr_min:
                    passed_rows += 1
                    snr_pass_values.append(row_snr_avg)
                    passed_heights.append(height)

        if passed_rows <= 0:
            return False, np.nan, "magasság: n/a"
        snr_mean = float(np.nanmean(np.asarray(snr_pass_values, dtype=float)))
        height_summary = self._format_dqc_height_summary(passed_heights)
        return True, snr_mean, height_summary

    def _run_dqc_filter(self, *_args):
        if not self._dqc_files:
            self._dqc_listbox.delete(0, tk.END)
            self._dqc_status_var.set("Válasszon ki egy megfelelő Data Quality mappát.")
            return

        try:
            height_low = float(self._dqc_height_low_var.get())
            height_high = float(self._dqc_height_high_var.get())
            zhzv_low = float(self._dqc_zhzv_low_var.get())
            zhzv_high = float(self._dqc_zhzv_high_var.get())
            rhohv_min = float(self._dqc_rhohv_min_var.get())
            snr_min = float(self._dqc_snr_min_var.get())
        except Exception:
            self._dqc_status_var.set("A szűrő paraméterei érvénytelenek.")
            return

        if height_low > height_high:
            height_low, height_high = height_high, height_low
        if zhzv_low > zhzv_high:
            zhzv_low, zhzv_high = zhzv_high, zhzv_low

        passed = []
        for radar_name, fname, file_path in self._dqc_files:
            ok, snr_mean, height_summary = self._dqc_file_passes(
                file_path, height_low, height_high, zhzv_low, zhzv_high, rhohv_min, snr_min
            )
            if ok:
                passed.append((radar_name, fname, snr_mean, height_summary))

        self._dqc_listbox.delete(0, tk.END)
        for radar_name, fname, snr_mean, height_summary in passed:
            self._dqc_listbox.insert(
                tk.END,
                f"{radar_name} | {fname} | átlag(SNRH,SNRV)={snr_mean:.2f} | {height_summary}"
            )

        self._dqc_status_var.set(
            f"Összes vizsgált fájl: {len(self._dqc_files)} | Átment: {len(passed)}"
        )

    def _collect_height_cloud_data(self, root_folder):
        heights = []
        zdr_values = []
        za_values = []
        for day_name in sorted(os.listdir(root_folder)):
            day_path = os.path.join(root_folder, day_name)
            if not os.path.isdir(day_path):
                continue
            for fname in sorted(os.listdir(day_path)):
                if not fname.lower().endswith(".txt"):
                    continue
                file_path = os.path.join(day_path, fname)
                if not os.path.isfile(file_path):
                    continue
                h, zdr, za = parse_file(file_path)
                if h.size == 0:
                    continue
                valid = ~np.isnan(h)
                if not valid.any():
                    continue
                heights.extend(h[valid].tolist())
                zdr_values.extend(zdr[valid].tolist())
                za_values.extend(za[valid].tolist())

        if not heights:
            return np.array([]), np.array([]), np.array([])

        return (np.array(heights, dtype=float),
                np.array(zdr_values, dtype=float),
                np.array(za_values, dtype=float))

    def _plot_height_folder_tab(self, heights, zdr_values, za_values):
        self._height_cloud_heights = heights
        self._height_cloud_zdr = zdr_values
        self._height_cloud_za = za_values
        self._height_avg_intervals = []
        self._height_interval_masks = []
        self._height_current_avg_maps = []

        valid_za = ~np.isnan(za_values)
        valid_zdr = ~np.isnan(zdr_values)

        if not valid_zdr.any() and not valid_za.any():
            self.fig6.clear()
            ax = self.fig6.add_subplot(111)
            ax.text(0.5, 0.5, "Nincs elérhető adat.",
                    transform=ax.transAxes, ha="center", va="center",
                    fontsize=11, color="#666")
            self.canvas6.draw()
            return

        if valid_za.any():
            zamin = np.floor(np.nanmin(za_values[valid_za]) / 5.0) * 5.0
            zamax = np.ceil(np.nanmax(za_values[valid_za]) / 5.0) * 5.0
        else:
            zamin, zamax = -10.0, 10.0

        bins = np.arange(zamin, zamax + 5.001, 5.0)
        if bins.size < 2:
            bins = np.array([zamin, zamin + 5.0])

        color_palette = ["#1f78b4", "#33a02c", "#e31a1c", "#ff7f00",
                         "#6a3d9a", "#b15928", "#a6cee3", "#b2df8a",
                         "#fb9a99", "#fdbf6f"]

        for idx in range(len(bins) - 1):
            lo, hi = bins[idx], bins[idx + 1]
            is_last = (idx == len(bins) - 2)
            mask = valid_za & (za_values >= lo) & (za_values <= hi if is_last else za_values < hi)
            color = color_palette[idx % len(color_palette)]

            self._height_interval_masks.append(mask)
            self._height_avg_intervals.append((lo, hi, color, {}))

        # only keep intervals that actually have data
        paired = [(m, t) for m, t in zip(self._height_interval_masks, self._height_avg_intervals)
                  if m.any()]
        if paired:
            self._height_interval_masks = [p[0] for p in paired]
            self._height_avg_intervals   = [p[1] for p in paired]

        self._height_selected_interval.set(-1)
        self._build_interval_buttons()
        self._draw_height_cloud(-1)

    def _build_interval_buttons(self):
        for w in self._height_interval_frame.winfo_children():
            w.destroy()
        tk.Label(self._height_interval_frame, text="Intervallum:",
                 bg="#f8f8f8", fg="#333", font=("Segoe UI", 9),
                 ).pack(side=tk.LEFT, padx=(0, 6))
        tk.Radiobutton(self._height_interval_frame, text="Összes",
                       variable=self._height_selected_interval, value=-1,
                       command=self._on_height_interval_select,
                       bg="#f8f8f8", activebackground="#e0e0e0",
                       font=("Segoe UI", 9), indicatoron=False,
                       padx=10, pady=3, relief=tk.RAISED, cursor="hand2",
                       ).pack(side=tk.LEFT, padx=2)
        for i, (lo, hi, color, _) in enumerate(self._height_avg_intervals):
            tk.Radiobutton(self._height_interval_frame,
                           text=f"{int(lo)}…{int(hi)} dBZ",
                           variable=self._height_selected_interval, value=i,
                           command=self._on_height_interval_select,
                           bg=color, fg="white", activebackground=color,
                           font=("Segoe UI", 9, "bold"), indicatoron=False,
                           padx=8, pady=3, relief=tk.RAISED, cursor="hand2",
                           ).pack(side=tk.LEFT, padx=2)

    def _on_height_interval_select(self):
        sel = self._height_selected_interval.get()
        self._draw_height_cloud(sel)
        self._attach_cursors()

    def _draw_height_cloud(self, sel_idx):
        if not hasattr(self, '_height_cloud_heights') or self._height_cloud_heights is None:
            return
        heights    = self._height_cloud_heights
        zdr_values = self._height_cloud_zdr
        za_values  = self._height_cloud_za
        intervals  = self._height_avg_intervals
        masks      = self._height_interval_masks

        filter_mask = self._height_filter_mask(zdr_values)
        self._height_current_avg_maps = []
        computed_intervals = []
        for i, (lo, hi, color, _unused) in enumerate(intervals):
            base_mask = masks[i] & filter_mask
            avg_map = self._build_height_average_map(heights, zdr_values, base_mask)
            self._height_current_avg_maps.append(avg_map)
            computed_intervals.append((lo, hi, color, avg_map))

        self.fig6.clear()
        ax1 = self.fig6.add_subplot(121)
        ax2 = self.fig6.add_subplot(122)
        self._height_cloud_ax1 = ax1

        valid_zdr = ~np.isnan(zdr_values)
        valid_za  = ~np.isnan(za_values)
        neutral   = "#bbbbbb"

        if sel_idx == -1:
            # neutral scatter + all interval averages on ZDR plot
            if valid_zdr.any():
                ax1.scatter(heights[valid_zdr], zdr_values[valid_zdr],
                            color=neutral, s=22, alpha=0.45, edgecolors="none", zorder=1)
            if valid_za.any():
                ax2.scatter(heights[valid_za], za_values[valid_za],
                            color=neutral, s=18, alpha=0.45, edgecolors="none", zorder=1)
            for lo, hi, color, avg_map in computed_intervals:
                ax2.axhspan(lo, hi, color=color, alpha=0.18, edgecolor="none")
                if avg_map:
                    avg_h = sorted(avg_map)
                    avg_z = [avg_map[h] for h in avg_h]
                    ax1.plot(avg_h, avg_z, color=color, lw=2.0, marker="D",
                             markersize=6, alpha=0.95, label=f"{int(lo)}…{int(hi)} dBZ")
        else:
            # dim background, highlight selected interval
            if valid_zdr.any():
                ax1.scatter(heights[valid_zdr], zdr_values[valid_zdr],
                            color="#dddddd", s=18, alpha=0.2, edgecolors="none", zorder=1)
            if valid_za.any():
                ax2.scatter(heights[valid_za], za_values[valid_za],
                            color="#dddddd", s=14, alpha=0.2, edgecolors="none", zorder=1)
            for i, (lo, hi, color, avg_map) in enumerate(computed_intervals):
                ax2.axhspan(lo, hi, color=color,
                            alpha=0.30 if i == sel_idx else 0.07, edgecolor="none")
                if i != sel_idx:
                    continue
                m_zdr = masks[i] & valid_zdr
                m_za  = masks[i] & valid_za
                if m_zdr.any():
                    ax1.scatter(heights[m_zdr], zdr_values[m_zdr],
                                color=color, s=44, alpha=0.92,
                                edgecolors="#222", linewidths=0.6, zorder=3,
                                label=f"ZDR – {int(lo)}…{int(hi)} dBZ")
                if m_za.any():
                    ax2.scatter(heights[m_za], za_values[m_za],
                                color=color, s=36, alpha=0.92,
                                edgecolors="#222", linewidths=0.6, zorder=3,
                                label=f"ZA – {int(lo)}…{int(hi)} dBZ")
                if avg_map:
                    avg_h = sorted(avg_map)
                    avg_z = [avg_map[h] for h in avg_h]
                    ax1.plot(avg_h, avg_z, color=color, lw=2.5, marker="D",
                             markersize=8, alpha=1.0, zorder=4,
                             label=f"Átlag ZDR – {int(lo)}…{int(hi)} dBZ")

        if self._height_stats_range is not None:
            low, high = self._height_stats_range
            ax1.axvspan(low, high, color="#ffdc7d", alpha=0.25, zorder=0)

        ax1.axhline(0, color="#999", lw=0.8, ls="--")
        ax1.set_xlabel("Magasság (km)", fontsize=11)
        ax1.set_ylabel("ZDR", fontsize=11)
        ax1.set_title("Magasság – ZDR", fontsize=12, fontweight="bold")
        ax1.grid(True, linestyle="--", alpha=0.4)
        if ax1.get_legend_handles_labels()[0]:
            ax1.legend(fontsize=8, loc="upper right")

        ax2.set_xlabel("Magasság (km)", fontsize=11)
        ax2.set_ylabel("ZA (reflektivitás)", fontsize=11)
        ax2.set_title("Magasság – ZA", fontsize=12, fontweight="bold")
        ax2.grid(True, linestyle="--", alpha=0.4)
        if ax2.get_legend_handles_labels()[0]:
            ax2.legend(fontsize=8, loc="upper right")

        if getattr(self, '_height_stats_enabled', None) is not None and self._height_stats_enabled.get():
            stats_text = self._height_stats_var.get() if getattr(self, '_height_stats_var', None) is not None else ""
            if stats_text:
                ax1.text(
                    0.99, 0.01, stats_text,
                    transform=ax1.transAxes,
                    ha="right", va="bottom",
                    fontsize=9,
                    color="#1a1a1a",
                    bbox=dict(boxstyle="round,pad=0.35", facecolor="white", alpha=0.7, edgecolor="none")
                )

        self.fig6.tight_layout(pad=2.0)
        self.canvas6.draw()

    def _height_cloud_hover(self, x, y):
        if not hasattr(self, '_height_current_avg_maps') or not self._height_current_avg_maps:
            return None
        sel = getattr(self, '_height_selected_interval', None)
        sel_idx = sel.get() if sel is not None else -1
        tol = 0.025
        if sel_idx >= 0:
            lo, hi, _, _ = self._height_avg_intervals[sel_idx]
            avg_map = self._height_current_avg_maps[sel_idx]
            if avg_map and x is not None:
                closest = min(avg_map, key=lambda h: abs(h - x))
                if abs(closest - x) <= tol:
                    return (f"{int(lo)}…{int(hi)} dBZ  →  "
                            f"Átlag ZDR: {avg_map[closest]:.4f} @ {closest:.3f} km")
            return None

        for i, (lo, hi, _, _unused) in enumerate(self._height_avg_intervals):
            avg_map = self._height_current_avg_maps[i]
            if not avg_map or x is None:
                continue
            closest = min(avg_map, key=lambda h: abs(h - x))
            if abs(closest - x) <= tol:
                return (f"{int(lo)}…{int(hi)} dBZ  →  "
                        f"Átlag ZDR: {avg_map[closest]:.4f} @ {closest:.3f} km")
        return None

    def _build_height_average_map(self, heights, zdr_values, base_mask):
        average_map = {}
        valid_mask = base_mask & ~np.isnan(heights) & ~np.isnan(zdr_values)
        if not valid_mask.any():
            return average_map
        selected_heights = heights[valid_mask]
        selected_zdr = zdr_values[valid_mask]
        for height in np.unique(selected_heights):
            values_at_height = selected_zdr[selected_heights == height]
            if values_at_height.size:
                average_map[float(height)] = float(np.mean(values_at_height))
        return average_map
    def _collect_month_zdr(self, root_folder):
        dates = []
        values = []
        samples = []
        for day_name in sorted(os.listdir(root_folder)):
            day_path = os.path.join(root_folder, day_name)
            if not os.path.isdir(day_path):
                continue
            for fname in sorted(os.listdir(day_path)):
                if not fname.lower().endswith(".txt"):
                    continue
                file_path = os.path.join(day_path, fname)
                if not os.path.isfile(file_path):
                    continue
                dt = self._parse_datetime_from_filename(fname)
                if dt is None:
                    continue
                _, zdr_vals, _ = parse_file(file_path)
                if zdr_vals.size == 0:
                    continue
                valid_zdr = zdr_vals[np.isfinite(zdr_vals)]
                if valid_zdr.size == 0:
                    continue
                mean_zdr = np.mean(valid_zdr)
                dates.append(dt)
                values.append(mean_zdr)
                samples.append(valid_zdr)
        return dates, np.array(values, dtype=float), samples

    @staticmethod
    def _calculate_month_zdr_stats(values):
        values = np.asarray(values, dtype=float)
        values = values[np.isfinite(values)]
        if values.size == 0:
            return 0, np.nan, np.nan
        mean = float(np.mean(values))
        std = float(np.std(values, ddof=1)) if values.size > 1 else np.nan
        return int(values.size), mean, std

    def _densest_cluster_mask(self, x_vals, y_vals):
        if x_vals.size < 3:
            return np.ones_like(x_vals, dtype=bool)
        H, x_edges, y_edges = np.histogram2d(x_vals, y_vals, bins=(25, 15))
        max_count = int(H.max())
        if max_count <= 1:
            return np.ones_like(x_vals, dtype=bool)

        threshold = max(1, max_count // 2)
        mask = np.zeros_like(x_vals, dtype=bool)
        for i in range(H.shape[0]):
            for j in range(H.shape[1]):
                if H[i, j] >= threshold:
                    in_x = (x_vals >= x_edges[i]) & (x_vals < x_edges[i + 1])
                    in_y = (y_vals >= y_edges[j]) & (y_vals < y_edges[j + 1])
                    mask |= (in_x & in_y)

        if mask.sum() < 2:
            return np.ones_like(x_vals, dtype=bool)
        return mask

    def _compute_piecewise_breaks(self, cluster_x, break_count):
        cluster_x = np.asarray(cluster_x, dtype=float)
        if break_count <= 0 or cluster_x.size < 3:
            return np.array([], dtype=float)

        max_breaks = max(1, min(break_count, cluster_x.size - 2))
        q = np.linspace(0.0, 1.0, max_breaks + 2)[1:-1]
        breaks = np.quantile(cluster_x, q)
        breaks = np.unique(np.asarray(breaks, dtype=float))
        breaks = breaks[(breaks > cluster_x.min()) & (breaks < cluster_x.max())]
        if breaks.size == 0:
            return np.array([], dtype=float)
        return breaks

    def _plot_month_folder_tab(self, dates, zdr_values, zdr_samples):
        self._disconnect_month_break_events()
        self.fig5.clear()
        ax = self.fig5.add_subplot(111)

        x_dates = np.array(dates)
        y_values = np.array(zdr_values, dtype=float)
        sample_groups = [np.asarray(values, dtype=float) for values in zdr_samples]
        self._month_dates = x_dates
        self._month_zdr_values = y_values
        self._month_zdr_samples = sample_groups

        ax.scatter(x_dates, y_values, color="#0078d4", s=40, alpha=0.8,
                   edgecolors="#333", linewidths=0.4, label="ZDR pontok")

        if self._month_daily_avg_enabled.get():
            daily_dates, daily_means = self._build_month_daily_means(x_dates, y_values)
            if daily_dates:
                ax.plot(daily_dates, daily_means,
                        color="#2ca02c", linewidth=2.0, marker="o", markersize=5,
                        alpha=0.95, label="Napi átlag ZDR")
        ax.axhline(0, color="gray", linewidth=0.8, linestyle="--")

        if len(x_dates) >= 2:
            x_nums = mdates.date2num(x_dates)
            mask = self._densest_cluster_mask(x_nums, y_values)
            cluster_x = x_nums[mask]
            cluster_y = y_values[mask]
            self._month_cluster_x = cluster_x
            self._month_cluster_y = cluster_y

            if cluster_x.size >= 2:
                if self._month_break_enabled.get() and cluster_x.size >= 4:
                    break_count = max(1, int(self._month_break_count_var.get()))
                    if self._month_break_points and any(np.isfinite(self._month_break_points)):
                        breaks = np.array(sorted(float(v) for v in self._month_break_points), dtype=float)
                        breaks = breaks[(breaks > cluster_x.min()) & (breaks < cluster_x.max())]
                        breaks = np.unique(np.round(breaks, 10))
                        if breaks.size != break_count:
                            breaks = self._compute_piecewise_breaks(cluster_x, break_count)
                    else:
                        breaks = self._compute_piecewise_breaks(cluster_x, break_count)

                    if breaks.size > 0:
                        self._month_break_points = [float(v) for v in breaks]
                        segment_bounds = [float(cluster_x.min())]
                        segment_bounds.extend(float(v) for v in breaks)
                        segment_bounds.append(float(cluster_x.max()))

                        for idx in range(len(segment_bounds) - 1):
                            left = segment_bounds[idx]
                            right = segment_bounds[idx + 1]
                            if right <= left:
                                continue
                            seg_mask = (cluster_x >= left) & (cluster_x <= right)
                            if np.count_nonzero(seg_mask) < 2:
                                continue
                            x_seg = cluster_x[seg_mask]
                            y_seg = cluster_y[seg_mask]
                            x_line = np.linspace(float(left), float(right), 80)

                            if self._month_break_mean_mode.get():
                                mean_value = float(np.nanmean(y_seg))
                                y_line = np.full_like(x_line, mean_value, dtype=float)
                                ax.plot(mdates.num2date(x_line), y_line,
                                        color="red", linewidth=2.2,
                                        alpha=0.9,
                                        label=f"Töréspontos átlag {idx + 1}. szakasz")
                            else:
                                y_line, y_low, y_high = self._compute_regression_and_ci(x_seg, y_seg, x_line)
                                ax.plot(mdates.num2date(x_line), y_line,
                                        color="red", linewidth=2.2,
                                        alpha=0.9,
                                        label=f"Illesztés {idx + 1}. szakasz")
                                if y_low is not None and y_high is not None:
                                    ax.fill_between(mdates.num2date(x_line), y_low, y_high,
                                                    color="red", alpha=0.12, edgecolor="none")

                        for b in breaks:
                            ax.axvline(float(b), color="black", linewidth=1.4, linestyle="--",
                                       alpha=0.9)
                        self._month_break_ax = ax
                        self._month_break_xlim = (float(cluster_x.min()), float(cluster_x.max()))
                    else:
                        slope, intercept = np.polyfit(cluster_x, cluster_y, 1)
                        x_line = np.linspace(cluster_x.min(), cluster_x.max(), 120)
                        y_line, y_low, y_high = self._compute_regression_and_ci(cluster_x, cluster_y, x_line)
                        ax.plot(mdates.num2date(x_line), y_line,
                                color="red", linewidth=2.2,
                                label="Legjobban illeszkedő piros vonal")
                        if y_low is not None:
                            ax.fill_between(mdates.num2date(x_line), y_low, y_high,
                                            color="red", alpha=0.18, edgecolor="none")
                        self._month_break_points = []
                else:
                    slope, intercept = np.polyfit(cluster_x, cluster_y, 1)
                    x_line = np.linspace(cluster_x.min(), cluster_x.max(), 120)
                    y_line, y_low, y_high = self._compute_regression_and_ci(cluster_x, cluster_y, x_line)
                    ax.plot(mdates.num2date(x_line), y_line,
                            color="red", linewidth=2.2,
                            label="Legjobban illeszkedő piros vonal")
                    if y_low is not None:
                        ax.fill_between(mdates.num2date(x_line), y_low, y_high,
                                        color="red", alpha=0.18, edgecolor="none")
                    self._month_break_points = []
        else:
            self._month_cluster_x = np.array([])
            self._month_cluster_y = np.array([])
            self._month_break_points = []

        ax.set_title("Havi ZDR pontfelhő", fontsize=13, fontweight="bold")

        x_nums_all = mdates.date2num(x_dates) if x_dates.size else np.array([])
        mean_label = ""
        if self._month_break_enabled.get() and self._month_break_points and x_nums_all.size:
            break_count_labels = []
            breaks = sorted(float(value) for value in self._month_break_points)
            for idx, breakpoint in enumerate(breaks):
                before_samples = np.concatenate([
                    sample_groups[i] for i in range(len(sample_groups))
                    if np.isfinite(x_nums_all[i]) and x_nums_all[i] <= breakpoint
                ]) if sample_groups else np.array([], dtype=float)
                after_samples = np.concatenate([
                    sample_groups[i] for i in range(len(sample_groups))
                    if np.isfinite(x_nums_all[i]) and x_nums_all[i] > breakpoint
                ]) if sample_groups else np.array([], dtype=float)
                before_count, before_mean, before_std = self._calculate_month_zdr_stats(
                    before_samples)
                after_count, after_mean, after_std = self._calculate_month_zdr_stats(
                    after_samples)
                prefix = f"{idx + 1}. " if len(breaks) > 1 else ""
                before_mean_text = f"{before_mean:.4f}" if np.isfinite(before_mean) else "n/a"
                before_std_text = f"{before_std:.4f}" if np.isfinite(before_std) else "n/a"
                after_mean_text = f"{after_mean:.4f}" if np.isfinite(after_mean) else "n/a"
                after_std_text = f"{after_std:.4f}" if np.isfinite(after_std) else "n/a"
                break_count_labels.append(
                    f"{prefix}előtte: N = {before_count}, nyers átlag = {before_mean_text}, "
                    f"s = {before_std_text} dB  |  utána: N = {after_count}, "
                    f"nyers átlag = {after_mean_text}, s = {after_std_text} dB")
            mean_label = "\n".join(break_count_labels)
        else:
            _, mean_total, _ = self._calculate_month_zdr_stats(
                np.concatenate(sample_groups) if sample_groups else np.array([]))
            if np.isfinite(mean_total):
                mean_label = f"Nyers ZDR átlag (teljes intervallum): {mean_total:.4f}"

        ax.text(0.01, 0.98, mean_label,
            transform=ax.transAxes,
            ha="left", va="top",
            fontsize=10, fontweight="bold", color="#1a5276",
            bbox=dict(boxstyle="round,pad=0.25", fc="#f4fbff", ec="#8ab6d6", alpha=0.9))

        ax.set_xlabel("Dátum", fontsize=11)
        ax.set_ylabel("Átlagos ZDR", fontsize=11)
        max_abs = max(8.0, np.nanmax(np.abs(y_values))) if y_values.size else 8.0
        ax.set_ylim(-max_abs, max_abs)
        ax.grid(True, linestyle="--", alpha=0.5)
        ax.xaxis.set_major_locator(mdates.AutoDateLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%m-%d"))
        self.fig5.autofmt_xdate(rotation=30)
        ax.legend(fontsize=9)
        self.fig5.tight_layout()
        self.canvas5.draw()
        self._bind_month_break_events()

    def _build_month_daily_means(self, x_dates, y_values):
        daily_totals = {}
        daily_counts = {}
        for dt, value in zip(x_dates, y_values):
            if np.isnan(value):
                continue
            day_key = dt.date()
            daily_totals[day_key] = daily_totals.get(day_key, 0.0) + float(value)
            daily_counts[day_key] = daily_counts.get(day_key, 0) + 1

        daily_dates = []
        daily_means = []
        for day_key in sorted(daily_totals.keys()):
            count = daily_counts.get(day_key, 0)
            if count <= 0:
                continue
            daily_dates.append(datetime.combine(day_key, datetime.min.time()))
            daily_means.append(daily_totals[day_key] / count)
        return daily_dates, daily_means

    def _on_month_ci_toggle(self):
        if hasattr(self, "_month_folder") and self._month_folder:
            self._plot_month_folder_tab(
                getattr(self, "_month_dates", []),
                getattr(self, "_month_zdr_values", np.array([])),
                getattr(self, "_month_zdr_samples", []),
            )
            self._attach_cursors()

    def _t_critical_value(self, df):
        if df <= 0:
            return 2.0
        critical = {
            1: 12.706,
            2: 4.303,
            3: 3.182,
            4: 2.776,
            5: 2.571,
            6: 2.447,
            7: 2.365,
            8: 2.306,
            9: 2.262,
            10: 2.228,
            11: 2.201,
            12: 2.179,
            13: 2.160,
            14: 2.145,
            15: 2.131,
            16: 2.120,
            17: 2.110,
            18: 2.101,
            19: 2.093,
            20: 2.086,
            21: 2.080,
            22: 2.074,
            23: 2.069,
            24: 2.064,
            25: 2.060,
            26: 2.056,
            27: 2.052,
            28: 2.048,
            29: 2.045,
            30: 2.042,
        }
        return critical.get(df, 1.96)

    def _compute_regression_and_ci(self, x_vals, y_vals, x_line):
        if not self._month_ci_enabled.get():
            slope, intercept = np.polyfit(x_vals, y_vals, 1)
            return slope * x_line + intercept, None, None

        n = x_vals.size
        if n < 3:
            slope, intercept = np.polyfit(x_vals, y_vals, 1)
            return slope * x_line + intercept, None, None

        slope, intercept = np.polyfit(x_vals, y_vals, 1)
        y_pred = slope * x_vals + intercept
        df = n - 2
        residuals = y_vals - y_pred
        s_err = np.sqrt(np.sum(residuals ** 2) / df) if df > 0 else 0.0
        x_mean = np.mean(x_vals)
        ssx = np.sum((x_vals - x_mean) ** 2)
        if ssx <= 0.0 or df <= 0:
            return slope * x_line + intercept, None, None

        t_val = self._t_critical_value(df)
        x_term = 1.0 / n + ((x_line - x_mean) ** 2) / ssx
        delta = t_val * s_err * np.sqrt(x_term)
        y_line = slope * x_line + intercept
        return y_line, y_line - delta, y_line + delta

    # ── Fájl megnyitás – 1. állomás ──────────────────────────────────────────
    def open_file(self):
        path = filedialog.askopenfilename(
            title="Válasszon .txt adatfájlt (1. állomás)",
            filetypes=[("Szöveges fájlok", "*.txt"), ("Minden fájl", "*.*")],
        )
        if not path:
            return
        try:
            heights, zdr, zatlag = parse_file(path)
            tab2_heights, tab2_series, tab2_mode = parse_file_tab2(path)
        except Exception as exc:
            messagebox.showerror("Olvasási hiba", f"Nem sikerült beolvasni a fájlt:\n{exc}")
            return
        if len(heights) == 0:
            messagebox.showwarning("Üres adatkészlet",
                                   "A fájlban nem találhatók érvényes numerikus sorok.")
            return

        self.heights, self.zdr, self.zatlag = heights, zdr, zatlag
        self.tab2_heights, self.tab2_series = tab2_heights, tab2_series
        self._tab2_mode = tab2_mode
        self._configure_tab2_mode(tab2_mode)
        self.cmp_heights1 = tab2_heights.copy()
        self.cmp_series1 = dict(tab2_series)
        self.cmp_mode1 = tab2_mode
        self._fullpath = path
        basename = os.path.basename(path)
        self.filepath.set(basename)
        self.status.set(
            f"Betöltve: {basename}  |  {len(heights)} érvényes sor  |  "
            f"Magasság: {heights.min():.2f} – {heights.max():.2f} km"
        )
        decoded = decode_filename(basename)
        self.file_info.set(f"Állomás / Időpont:   {decoded}" if decoded else "")
        self._cmp_info1.set(f"1. fájl: {decoded}" if decoded else f"1. fájl: {basename}")

        img_url, page_url = build_radar_url(basename)
        if img_url:
            self._radar_url = img_url
            self._radar_url_var.set(page_url)
        else:
            self._radar_url = None
            self._radar_url_var.set("– (fájlnév nem azonosítható)")

        self._plot_tab2()
        self._configure_compare_mode()
        self._plot_comparison()
        self._plot_diff_tab()
        self._plot_delta_zero_tab()
        self._attach_cursors()

        if img_url:
            self._load_radar_image()

    # ── Fájl megnyitás – 2. állomás ──────────────────────────────────────────
    def open_file2(self):
        path = filedialog.askopenfilename(
            title="Válasszon .txt adatfájlt (2. állomás)",
            filetypes=[("Szöveges fájlok", "*.txt"), ("Minden fájl", "*.*")],
        )
        if not path:
            return
        try:
            heights, zdr, zatlag = parse_file(path)
            cmp_heights, cmp_series, cmp_mode = parse_file_tab2(path)
        except Exception as exc:
            messagebox.showerror("Olvasási hiba", f"Nem sikerült beolvasni a fájlt:\n{exc}")
            return
        if len(heights) == 0:
            messagebox.showwarning("Üres adatkészlet",
                                   "A fájlban nem találhatók érvényes numerikus sorok.")
            return

        self.heights2, self.zdr2, self.zatlag2 = heights, zdr, zatlag
        self.cmp_heights2 = cmp_heights.copy()
        self.cmp_series2 = dict(cmp_series)
        self.cmp_mode2 = cmp_mode
        basename = os.path.basename(path)
        self._filepath2_var.set(basename)
        decoded = decode_filename(basename)
        self._cmp_info2.set(f"2. fájl: {decoded}" if decoded else f"2. fájl: {basename}")

        img_url, page_url = build_radar_url(basename)
        if img_url:
            self._radar_url2 = img_url
            self._radar_url_var2.set(page_url)
        else:
            self._radar_url2 = None
            self._radar_url_var2.set("– (fájlnév nem azonosítható)")

        self._configure_compare_mode()
        self._plot_comparison()
        self._plot_diff_tab()
        self._plot_delta_zero_tab()
        self._attach_cursors()

        if img_url:
            self._load_radar_image2()

    def _on_delta_limit_change(self, value):
        try:
            self._delta_limit_var.set(float(value))
        except ValueError:
            return
        self._plot_delta_zero_tab()
        self._attach_cursors()

    def _configure_tab2_mode(self, mode):
        self._tab2_series_order = list(self._tab2_series_orders.get(mode, self._tab2_series_orders["12"]))
        defaults = self._tab2_defaults.get(mode, self._tab2_defaults["12"])
        self._tab2_visible_vars = {
            name: tk.BooleanVar(value=defaults.get(name, False))
            for name in self._tab2_series_order
        }
        self._render_tab2_controls()

    def _cmp_series_for_mode(self, mode):
        if mode is None:
            return []
        return list(self._tab2_series_orders.get(mode, []))

    def _configure_compare_mode(self):
        available1 = self._cmp_series_for_mode(self.cmp_mode1)
        available2 = self._cmp_series_for_mode(self.cmp_mode2)
        if available1 and available2:
            common = [name for name in available1 if name in set(available2)]
        else:
            common = available1 or available2

        self._cmp_series_order = list(common)
        defaults = {name: False for name in self._cmp_series_order}
        if self.cmp_mode1 == self.cmp_mode2 == "3" and "Zátlag" in defaults:
            defaults["Zátlag"] = True
        if "ZDR" in defaults:
            defaults["ZDR"] = True

        self._cmp_visible_vars = {
            name: tk.BooleanVar(value=defaults.get(name, False))
            for name in self._cmp_series_order
        }
        self._render_compare_controls()

    def _render_compare_controls(self):
        if not self._cmp_controls_strip:
            return
        for widget in self._cmp_controls_strip.winfo_children():
            widget.destroy()

        if not self._cmp_series_order:
            tk.Label(
                self._cmp_controls_strip,
                text="Adatsorok: nincs közös adat",
                bg="#f7f7f7", fg="#666", font=("Segoe UI", 9),
            ).pack(side=tk.LEFT)
            return

        row_canvas = tk.Canvas(self._cmp_controls_strip, bg="#f7f7f7", height=38, highlightthickness=0)
        row_canvas.pack(side=tk.LEFT, fill=tk.X, expand=True)
        row_scroll = tk.Scrollbar(self._cmp_controls_strip, orient=tk.HORIZONTAL, command=row_canvas.xview)
        row_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        row_canvas.configure(xscrollcommand=row_scroll.set)

        checks = tk.Frame(row_canvas, bg="#f7f7f7")
        row_canvas.create_window((0, 0), window=checks, anchor="nw")

        for name in self._cmp_series_order:
            cb = tk.Checkbutton(
                checks,
                text=name,
                variable=self._cmp_visible_vars[name],
                command=self._plot_comparison,
                bg="#f7f7f7",
                fg="#222",
                selectcolor="#dff0d8",
                activebackground="#f7f7f7",
                font=("Segoe UI", 9),
                relief=tk.FLAT,
                padx=5,
                pady=2,
                cursor="hand2",
                indicatoron=False,
                bd=1,
            )
            cb.pack(side=tk.LEFT, padx=(0, 4))

        def _sync_cmp_strip(event=None):
            row_canvas.configure(scrollregion=row_canvas.bbox("all"))

        checks.bind("<Configure>", _sync_cmp_strip)
        row_canvas.bind("<Configure>", _sync_cmp_strip)

    def _active_compare_series(self):
        return [name for name in self._cmp_series_order if self._cmp_visible_vars.get(name) and self._cmp_visible_vars[name].get()]

    def _render_tab2_controls(self):
        if not hasattr(self, "_tab2_controls_strip"):
            return
        for widget in self._tab2_controls_strip.winfo_children():
            widget.destroy()

        if not self._tab2_series_order:
            tk.Label(
                self._tab2_controls_strip,
                text="Adatsorok: nincs betöltött adat",
                bg="#f7f7f7", fg="#666", font=("Segoe UI", 9),
            ).pack(side=tk.LEFT)
            return

        row_canvas = tk.Canvas(self._tab2_controls_strip, bg="#f7f7f7", height=38, highlightthickness=0)
        row_canvas.pack(side=tk.LEFT, fill=tk.X, expand=True)
        row_scroll = tk.Scrollbar(self._tab2_controls_strip, orient=tk.HORIZONTAL, command=row_canvas.xview)
        row_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        row_canvas.configure(xscrollcommand=row_scroll.set)

        checks = tk.Frame(row_canvas, bg="#f7f7f7")
        row_canvas.create_window((0, 0), window=checks, anchor="nw")

        for name in self._tab2_series_order:
            cb = tk.Checkbutton(
                checks,
                text=name,
                variable=self._tab2_visible_vars[name],
                command=self._on_tab2_series_toggle,
                bg="#f7f7f7",
                fg="#222",
                selectcolor="#dff0d8",
                activebackground="#f7f7f7",
                font=("Segoe UI", 9),
                relief=tk.FLAT,
                padx=5,
                pady=2,
                cursor="hand2",
                indicatoron=False,
                bd=1,
            )
            cb.pack(side=tk.LEFT, padx=(0, 4))

        def _sync_tab2_strip(event=None):
            row_canvas.configure(scrollregion=row_canvas.bbox("all"))

        checks.bind("<Configure>", _sync_tab2_strip)
        row_canvas.bind("<Configure>", _sync_tab2_strip)

    def _on_tab2_series_toggle(self):
        if self.tab2_heights is None:
            return
        self._plot_tab2()
        self._attach_cursors()

    def _plot_tab2(self):
        self.fig2.clear()
        ax = self.fig2.add_subplot(111)

        heights = self.tab2_heights if self.tab2_heights.size else self.heights
        series = {name: np.array([]) for name in self._tab2_series_order}
        if self.tab2_series:
            for name in self._tab2_series_order:
                if name in self.tab2_series:
                    series[name] = self.tab2_series[name]
        elif self.zdr is not None and self._tab2_mode == "3":
            series["ZDR"] = self.zdr
            series["Zátlag"] = self.zatlag if self.zatlag is not None else np.array([])

        plotted = False
        for name in self._tab2_series_order:
            if not self._tab2_visible_vars[name].get():
                continue
            vals = series.get(name, np.array([]))
            if heights is None or vals is None or len(heights) == 0 or len(vals) != len(heights):
                continue
            if np.all(np.isnan(vals)):
                continue
            ax.plot(
                heights, vals,
                color=self._tab2_palette.get(name, None), linewidth=1.4,
                marker=self._tab2_markers.get(name, "o"), markersize=2.7,
                label=name,
            )
            plotted = True

        if not plotted:
            ax.text(0.5, 0.5,
                    "Nincs aktív vagy megjeleníthető adatsor a 2. fülön.",
                    transform=ax.transAxes, ha="center", va="center",
                    fontsize=11, color="#666")

        ax.axhline(0, color="gray", linewidth=0.7, linestyle="--")
        ax.set_xlabel("Magasság (km)", fontsize=11)
        ax.set_ylabel("Érték", fontsize=11)
        if self._tab2_mode == "3":
            ax.set_title("Magasság – ZDR és Zátlag", fontsize=13, fontweight="bold")
        else:
            ax.set_title("Magasság – választható adatsorok", fontsize=13, fontweight="bold")
        if plotted:
            ax.legend(fontsize=9, ncol=2)
        ax.grid(True, linestyle="--", alpha=0.5)
        self.fig2.tight_layout()
        self.canvas2.draw()

    def _plot_diff_tab(self):
        self.fig3.clear()
        ax = self.fig3.add_subplot(111)

        h1  = self.heights  if self.heights  is not None else np.array([])
        z1  = self.zdr      if self.zdr      is not None else np.array([])
        za1 = self.zatlag   if self.zatlag   is not None else np.array([])
        h2  = self.heights2 if self.heights2 is not None else np.array([])
        z2  = self.zdr2     if self.zdr2     is not None else np.array([])
        za2 = self.zatlag2  if self.zatlag2 is not None else np.array([])

        if len(h1) == 0 or len(h2) == 0:
            ax.text(0.5, 0.5,
                    "Töltsön be mindkét állomás adatait a különbségek megjelenítéséhez.",
                    transform=ax.transAxes, ha="center", va="center",
                    fontsize=11, color="#666")
            ax.set_xlabel("Magasság (km)", fontsize=11)
            ax.set_ylabel("Abszolút különbség", fontsize=11)
            ax.set_title("ΔZDR & ΔZátlag", fontsize=13, fontweight="bold")
            ax.grid(True, linestyle="--", alpha=0.5)
            self.fig3.tight_layout()
            self.canvas3.draw()
            return

        def last_valid_height(heights_arr, zdr_arr, za_arr):
            valid = (~np.isnan(zdr_arr)) | (~np.isnan(za_arr))
            if not valid.any():
                return np.nan
            return float(np.max(heights_arr[valid]))

        last1 = last_valid_height(h1, z1, za1)
        last2 = last_valid_height(h2, z2, za2)
        if np.isnan(last1) or np.isnan(last2):
            ax.text(0.5, 0.5,
                    "Legalább egy állomásnál nincs érvényes ZDR vagy Zátlag adat.",
                    transform=ax.transAxes, ha="center", va="center",
                    fontsize=11, color="#666")
            ax.set_xlabel("Magasság (km)", fontsize=11)
            ax.set_ylabel("Abszolút különbség", fontsize=11)
            ax.set_title("ΔZDR & ΔZátlag", fontsize=13, fontweight="bold")
            ax.grid(True, linestyle="--", alpha=0.5)
            self.fig3.tight_layout()
            self.canvas3.draw()
            return

        max_height = max(last1, last2)
        heights = np.unique(np.concatenate((h1, h2)))
        heights = heights[heights <= max_height]

        def safe_interp(points, x_arr, y_arr):
            mask = ~np.isnan(x_arr) & ~np.isnan(y_arr)
            if mask.sum() < 2:
                return np.full_like(points, np.nan, dtype=float)
            x_clean = x_arr[mask]
            y_clean = y_arr[mask]
            order = np.argsort(x_clean)
            x_sorted = x_clean[order]
            y_sorted = y_clean[order]
            return np.interp(points, x_sorted, y_sorted, left=np.nan, right=np.nan)

        def explicit_nan_mask(points, x_arr, y_arr):
            mask = np.full(points.shape, True, dtype=bool)
            for i, val in enumerate(points):
                match = np.where(x_arr == val)[0]
                if match.size > 0:
                    mask[i] = ~np.isnan(y_arr[match]).any()
            return mask

        z1_interp  = safe_interp(heights, h1,  z1)
        z2_interp  = safe_interp(heights, h2,  z2)
        za1_interp = safe_interp(heights, h1,  za1)
        za2_interp = safe_interp(heights, h2,  za2)

        valid_zdr = (~np.isnan(z1_interp) & ~np.isnan(z2_interp)
                     & explicit_nan_mask(heights, h1, z1)
                     & explicit_nan_mask(heights, h2, z2))
        valid_za  = (~np.isnan(za1_interp) & ~np.isnan(za2_interp)
                     & explicit_nan_mask(heights, h1, za1)
                     & explicit_nan_mask(heights, h2, za2))
        plotted = False

        if valid_zdr.any():
            ax.plot(heights[valid_zdr], np.abs(z1_interp[valid_zdr] - z2_interp[valid_zdr]),
                    color="#0078d4", lw=1.8, marker="o", ms=3, label="|ΔZDR|")
            plotted = True
        if valid_za.any():
            ax.plot(heights[valid_za], np.abs(za1_interp[valid_za] - za2_interp[valid_za]),
                    color="#e05c00", lw=1.8, marker="s", ms=3, label="|ΔZátlag|")
            plotted = True
        if not plotted:
            ax.text(0.5, 0.5,
                    "Nincs számítható különbség a rendelkezésre álló adatból.",
                    transform=ax.transAxes, ha="center", va="center",
                    fontsize=11, color="#666")

        ax.set_xlabel("Magasság (km)", fontsize=11)
        ax.set_ylabel("Abszolút különbség", fontsize=11)
        ax.set_title("ΔZDR & ΔZátlag", fontsize=13, fontweight="bold")
        ax.legend(fontsize=10)
        ax.grid(True, linestyle="--", alpha=0.5)
        self.fig3.tight_layout()
        self.canvas3.draw()

    def _plot_delta_zero_tab(self):
        """Plot how much later measurement's ZDR moved closer/further to 0
        compared to the earlier measurement from filename timestamps.
        Also plot later ZA minus earlier ZA.
        """
        # ensure fig4 exists
        if not hasattr(self, 'fig4'):
            return
        self.fig4.clear()
        ax = self.fig4.add_subplot(111)

        b1 = self.filepath.get() if isinstance(self.filepath, tk.StringVar) else str(self.filepath)
        b2 = self._filepath2_var.get() if isinstance(self._filepath2_var, tk.StringVar) else str(self._filepath2_var)
        if not b1 or not b2 or b1.startswith("Nincs") or b2.startswith("Nincs"):
            ax.text(0.5, 0.5,
                    "Töltsön be mindkét állomás adatait a vizsgálathoz.",
                    transform=ax.transAxes, ha="center", va="center",
                    fontsize=11, color="#666")
            ax.set_xlabel("Magasság (km)", fontsize=11)
            ax.set_ylabel("Változás", fontsize=11)
            ax.set_title("Δ a 0-hoz (időbeli változás)", fontsize=13, fontweight="bold")
            self.fig4.tight_layout()
            self.canvas4.draw()
            return

        def parse_dt_from_name(name):
            base = os.path.splitext(name)[0]
            if len(base) < 15:
                return None
            digits = base[3:15]
            if not digits.isdigit():
                return None
            try:
                y = int(digits[0:4]); mo = int(digits[4:6]); d = int(digits[6:8])
                hh = int(digits[8:10]); mm = int(digits[10:12])
                return datetime(y, mo, d, hh, mm)
            except Exception:
                return None

        dt1 = parse_dt_from_name(b1)
        dt2 = parse_dt_from_name(b2)

        # Assign earlier and later
        if dt1 is not None and dt2 is not None:
            if dt1 <= dt2:
                h_early, z_early, za_early = (self.heights, self.zdr, self.zatlag)
                h_late, z_late, za_late = (self.heights2, self.zdr2, self.zatlag2)
                label_early = b1
                label_late = b2
            else:
                h_early, z_early, za_early = (self.heights2, self.zdr2, self.zatlag2)
                h_late, z_late, za_late = (self.heights, self.zdr, self.zatlag)
                label_early = b2
                label_late = b1
        else:
            # fallback: treat file1 as earlier, file2 as later
            h_early, z_early, za_early = (self.heights, self.zdr, self.zatlag)
            h_late, z_late, za_late = (self.heights2, self.zdr2, self.zatlag2)
            label_early = b1
            label_late = b2

        if h_early is None or h_late is None:
            if self._delta_limit_scale:
                self._delta_limit_scale.config(state=tk.DISABLED)
                self._delta_limit_label.set("Felső határ: – km")
            ax.text(0.5, 0.5,
                    "Mindkét állomás adata szükséges.",
                    transform=ax.transAxes, ha="center", va="center",
                    fontsize=11, color="#666")
            self.fig4.tight_layout()
            self.canvas4.draw()
            return

        tol = 1e-6
        h1r = np.round(h_early, 6)
        h2r = np.round(h_late, 6)
        common = np.intersect1d(h1r, h2r)
        if common.size == 0:
            if self._delta_limit_scale:
                self._delta_limit_scale.config(state=tk.DISABLED)
                self._delta_limit_label.set("Felső határ: – km")
            ax.text(0.5, 0.5,
                    "Nincs elég közös magassági pont a számításhoz.",
                    transform=ax.transAxes, ha="center", va="center",
                    fontsize=11, color="#666")
            self.fig4.tight_layout()
            self.canvas4.draw()
            return

        min_common = float(common.min())
        max_common = float(common.max())
        if self._delta_limit_scale:
            self._delta_limit_scale.config(from_=min_common, to=max_common,
                                           resolution=max(0.01, (max_common - min_common)/1000),
                                           state=tk.NORMAL)
            if self._delta_limit_var.get() < min_common or self._delta_limit_var.get() > max_common:
                self._delta_limit_var.set(max_common)
        else:
            self._delta_limit_var.set(max_common)
        threshold = float(self._delta_limit_var.get())
        self._delta_limit_label.set(f"Felső határ: {threshold:.3f} km")

        zdr_diffs = np.full(common.shape, np.nan, dtype=float)
        za_diffs = np.full(common.shape, np.nan, dtype=float)
        for idx, height in enumerate(common):
            i_e = np.where(np.isclose(h_early, height, atol=tol))[0]
            i_l = np.where(np.isclose(h_late, height, atol=tol))[0]
            if i_e.size == 0 or i_l.size == 0:
                continue
            ze = z_early[i_e[0]]
            zl = z_late[i_l[0]]
            if not np.isnan(ze) and not np.isnan(zl):
                zdr_diffs[idx] = abs(zl) - abs(ze)
            ae = za_early[i_e[0]]
            al = za_late[i_l[0]]
            if not np.isnan(ae) and not np.isnan(al):
                za_diffs[idx] = al - ae

        threshold_mask = common <= threshold
        filtered_common = common[threshold_mask]
        filtered_zdr_diffs = zdr_diffs[threshold_mask]
        filtered_za_diffs = za_diffs[threshold_mask]

        def build_from_common(heights, diffs):
            xdata = []
            ydata = []
            gap = False
            for h, d in zip(heights, diffs):
                if np.isnan(d):
                    if xdata:
                        gap = True
                    continue
                if gap:
                    xdata.append(np.nan); ydata.append(np.nan)
                    gap = False
                xdata.append(h); ydata.append(d)
            return np.array(xdata, dtype=float), np.array(ydata, dtype=float)

        xz_full, yz_full = build_from_common(common, zdr_diffs)
        ya_x_full, ya_y_full = build_from_common(common, za_diffs)

        threshold_mask = common <= threshold
        filtered_common = common[threshold_mask]
        filtered_zdr_diffs = zdr_diffs[threshold_mask]
        filtered_za_diffs = za_diffs[threshold_mask]

        xz, yz = build_from_common(filtered_common, filtered_zdr_diffs)
        ya_x, ya_y = build_from_common(filtered_common, filtered_za_diffs)

        threshold_line = ax.axvline(threshold, color="#cc0000", linewidth=1.2,
                                    linestyle="--", alpha=0.8,
                                    label="_threshold")

        def integrate_above_below(x_vals, y_vals):
            area_pos = 0.0
            area_neg = 0.0
            if x_vals.size <= 1:
                return area_pos, area_neg
            valid_idx = [i for i in range(x_vals.size)
                         if not np.isnan(x_vals[i]) and not np.isnan(y_vals[i])]
            for i in range(len(valid_idx) - 1):
                idx1, idx2 = valid_idx[i], valid_idx[i + 1]
                x1, x2 = x_vals[idx1], x_vals[idx2]
                y1, y2 = y_vals[idx1], y_vals[idx2]
                area_segment = (y1 + y2) / 2.0 * (x2 - x1)
                if y1 >= 0 and y2 >= 0:
                    area_pos += area_segment
                elif y1 <= 0 and y2 <= 0:
                    area_neg -= area_segment
                else:
                    if y1 * y2 < 0:
                        x_cross = x1 - y1 * (x2 - x1) / (y2 - y1)
                        area1 = (y1 + 0) / 2.0 * (x_cross - x1)
                        area2 = (0 + y2) / 2.0 * (x2 - x_cross)
                        if y1 > 0:
                            area_pos += area1
                            area_neg -= area2
                        else:
                            area_neg -= area1
                            area_pos += area2
            return area_pos, area_neg

        area_above, area_below = integrate_above_below(xz, yz)
        za_area_above, za_area_below = integrate_above_below(ya_x, ya_y)

        # Compute Pearson correlation for matched ZDR and ZA diff values on the full dataset
        corr_value = np.nan
        valid_corr = ~np.isnan(zdr_diffs) & ~np.isnan(za_diffs)
        if np.count_nonzero(valid_corr) > 1:
            corr_matrix = np.corrcoef(zdr_diffs[valid_corr], za_diffs[valid_corr])
            corr_value = float(corr_matrix[0, 1])

        if xz_full.size > 0:
            ax.plot(xz_full, yz_full, color="#0078d4", lw=1.2, marker="o", ms=3,
                    alpha=0.75, label=f"Abs(ZDR) change (teljes adat)")
        if ya_x_full.size > 0:
            ax.plot(ya_x_full, ya_y_full, color="#e05c00", lw=1.2, marker="s", ms=3,
                    alpha=0.75, label=f"ΔZátlag (teljes adat)")
        if xz.size > 0:
            ax.plot(xz, yz, color="#0078d4", lw=2.0, marker="o", ms=4,
                    label=f"Abs(ZDR) change (szűrt)")
        if ya_x.size > 0:
            ax.plot(ya_x, ya_y, color="#e05c00", lw=2.0, marker="s", ms=4,
                    label=f"ΔZátlag (szűrt)")

        ax.axhline(0, color="gray", linewidth=0.8, linestyle="--")
        ax.set_xlabel("Magasság (km)", fontsize=11)
        ax.set_ylabel("Változás", fontsize=11)
        ax.set_title(f"Δ a 0-hoz: {label_late} vs {label_early}", fontsize=12)
        
        # Add draggable, transparent text box with integrated areas and correlation
        corr_text = (
            f"Pearson korreláció: {corr_value:+.4f}" if not np.isnan(corr_value)
            else "Pearson korreláció: nem elég adat"
        )
        info_text = (
            f"ZDR növekedés (felül): {area_above:+.4f}\n"
            f"ZDR csökkenés (alatt): {area_below:+.4f}\n"
            f"ZA növekedés (felül): {za_area_above:+.4f}\n"
            f"ZA csökkenés (alatt): {za_area_below:+.4f}\n"
            f"{corr_text}"
        )
        
        # Create draggable annotation box with high transparency
        ann = ax.annotate(info_text, xy=(0.02, 0.98), xycoords='axes fraction',
                          fontsize=10, family='monospace',
                          bbox=dict(boxstyle='round,pad=0.5', fc='lightyellow', 
                                    ec='#aaa', alpha=0.2),
                          ha='left', va='top')
        ann.draggable(True)
        
        ax.legend(fontsize=10)
        ax.grid(True, linestyle="--", alpha=0.5)
        self.fig4.tight_layout()
        self.canvas4.draw()

    def _plot_comparison(self):
        """Összehasonlítás: két subplot egymás alatt, közös X (magasság) tengely."""
        self.figcmp.clear()

        active_series = self._active_compare_series()
        h1 = self.cmp_heights1 if self.cmp_heights1 is not None else np.array([])
        h2 = self.cmp_heights2 if self.cmp_heights2 is not None else np.array([])
        series1 = self.cmp_series1 or {}
        series2 = self.cmp_series2 or {}

        ax1 = self.figcmp.add_subplot(211)
        ax2 = self.figcmp.add_subplot(212, sharex=ax1)

        # ─ 1. subplot
        title1 = self._cmp_info1.get().replace("1. fájl: ", "")
        if len(h1) > 0:
            for series_name in active_series:
                vals = series1.get(series_name, np.array([]))
                if len(vals) != len(h1) or np.all(np.isnan(vals)):
                    continue
                ax1.plot(h1, vals,
                         color=self._cmp_palette.get(series_name), lw=1.8,
                         marker=self._cmp_markers.get(series_name, "o"), ms=3,
                         label=series_name)
        else:
            ax1.text(0.5, 0.5, "Nincs betöltött adat – nyisson meg 1. fájlt",
                     transform=ax1.transAxes, ha="center", va="center",
                     fontsize=11, color="#aaa")
        ax1.axhline(0, color="gray", lw=0.7, ls="--")
        ax1.set_ylabel("Érték", fontsize=10)
        ax1.set_title(f"1. állomás:  {title1}", fontsize=11, fontweight="bold", color="#1a5276")
        if any(not str(line.get_label()).startswith("_") for line in ax1.lines):
            ax1.legend(fontsize=9)
        ax1.grid(True, ls="--", alpha=0.5)

        # ─ 2. subplot
        title2 = self._cmp_info2.get().replace("2. fájl: ", "")
        if len(h2) > 0:
            for series_name in active_series:
                vals = series2.get(series_name, np.array([]))
                if len(vals) != len(h2) or np.all(np.isnan(vals)):
                    continue
                ax2.plot(h2, vals,
                         color=self._cmp_palette.get(series_name), lw=1.8,
                         marker=self._cmp_markers.get(series_name, "o"), ms=3,
                         label=series_name)
        else:
            ax2.text(0.5, 0.5, "Töltsön be 2. állomás adatát a zöld gombbal",
                     transform=ax2.transAxes, ha="center", va="center",
                     fontsize=11, color="#aaa")
        ax2.axhline(0, color="gray", lw=0.7, ls="--")
        ax2.set_xlabel("Magasság (km)", fontsize=10)
        ax2.set_ylabel("Érték", fontsize=10)
        ax2.set_title(f"2. állomás:  {title2}", fontsize=11, fontweight="bold", color="#196f3d")
        if any(not str(line.get_label()).startswith("_") for line in ax2.lines):
            ax2.legend(fontsize=9)
        ax2.grid(True, ls="--", alpha=0.5)

        self.figcmp.tight_layout(h_pad=3.5)
        self.canvascmp.draw()

        # Összehasonlítás kurzor csatolása
        self._cmp_cursor = ComparisonCrosshair(ax1, ax2, self.canvascmp,
                                               self._cmp_coord_var)
        self._cmp_cursor.toggle(self._cmp_cursor_on.get())


if __name__ == "__main__":
    app = RadarPlotterApp()
    app.mainloop()
