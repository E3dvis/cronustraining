import sys
import os
import json
import time
import random
import subprocess
from datetime import datetime
import requests
import csv
import io

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QLineEdit, QFrame, QScrollArea, QSpinBox, QGroupBox, QCheckBox,
    QTextEdit, QDialog, QFileDialog, QMessageBox, QMenu, QToolTip,
    QProgressBar, QComboBox, QRadioButton
)
from PyQt6.QtCore import (
    Qt, QTimer, QThread, pyqtSignal, QPoint, QEasingCurve, QPropertyAnimation,
    QRectF, pyqtProperty
)
from PyQt6.QtGui import (
    QFont, QPixmap, QPainter, QColor, QPen, QBrush, QLinearGradient
)

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import matplotlib.pyplot as plt
from matplotlib.figure import Figure

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors

plt.rcParams['font.family'] = 'DejaVu Sans'

API_BASE = "http://127.0.0.1:35100/v0/Cronus"

DEFAULT_PARAMS = {
    "test_min": None,
    "test_max": None,
    "wait_time": 3.0,
    "cycles": 100,
    "measure_power_curve": False
}

CONFIG_FILENAME = "cronus_app_config.json"
CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), CONFIG_FILENAME)
LOG_BASE = "D:/CronusTestingLogs"

BASE_ZONE_DEFS = [
    {"fixed_min": 670.0,   "fixed_max": 800.0,   "color": "#9ec5fe"},
    {"fixed_min": 800.1,   "fixed_max": 960.0,   "color": "#d6bcfa"},
    {"fixed_min": 940.0,   "fixed_max": 1100.0,  "color": "#fecaca"},
    {"fixed_min": 1100.1,  "fixed_max": 1320.0,  "color": "#c8facc"},
]

def ensure_log_dirs():
    os.makedirs(LOG_BASE, exist_ok=True)
    return LOG_BASE

def set_log_dir(new_dir: str):
    global LOG_BASE
    if new_dir:
        try:
            os.makedirs(new_dir, exist_ok=True)
            LOG_BASE = new_dir
        except Exception as e:
            print(f"Failed to create log directory {new_dir}: {e}")
    ensure_log_dirs()

def safe_get_json(url, timeout=2):
    try:
        r = requests.get(url, timeout=timeout)
        r.raise_for_status()
        return r.json()
    except:
        return None

def safe_put_json(url, payload, timeout=3):
    try:
        r = requests.put(url, json=payload, timeout=timeout)
        r.raise_for_status()
        return r.json()
    except:
        return None

def check_connection(channel: int):
    try:
        response = safe_get_json(f"{API_BASE}/Ch{channel}/Status")
        return response is not None and response.get("OK", False)
    except:
        return False

def fetch_device_range(channel: int):
    rng = safe_get_json(f"{API_BASE}/Ch{channel}/WavelengthRange")
    if rng and rng.get("OK") and not rng.get("IsEmpty"):
        try:
            return (float(rng.get("Min")), float(rng.get("Max")))
        except:
            return (None, None)
    return (None, None)

def load_config():
    global LOG_BASE
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            data = {}
    else:
        data = {}
    LOG_BASE = data.get("log_dir", LOG_BASE)
    ensure_log_dirs()

    def merge_channel(raw):
        merged = {}
        merged["test_min"] = raw.get("test_min", None)
        merged["test_max"] = raw.get("test_max", None)
        merged["wait_time"] = raw.get("wait_time", DEFAULT_PARAMS["wait_time"])
        merged["cycles"] = raw.get("cycles", DEFAULT_PARAMS["cycles"])
        merged["measure_power_curve"] = raw.get("measure_power_curve", DEFAULT_PARAMS["measure_power_curve"])
        return merged

    # Zones
    zones_cfg = data.get("zones")
    def default_zone_list():
        return [
            {"name": "VIS1", "enabled": True, "min": BASE_ZONE_DEFS[0]["fixed_min"], "max": BASE_ZONE_DEFS[0]["fixed_max"]},
            {"name": "VIS2", "enabled": True, "min": BASE_ZONE_DEFS[1]["fixed_min"], "max": BASE_ZONE_DEFS[1]["fixed_max"]},
            {"name": "IR1",  "enabled": True, "min": BASE_ZONE_DEFS[2]["fixed_min"], "max": BASE_ZONE_DEFS[2]["fixed_max"]},
            {"name": "IR2",  "enabled": True, "min": BASE_ZONE_DEFS[3]["fixed_min"], "max": BASE_ZONE_DEFS[3]["fixed_max"]},
        ]
    if not isinstance(zones_cfg, list) or len(zones_cfg) != 4:
        zones_cfg = default_zone_list()
    else:
        for i, z in enumerate(zones_cfg):
            if "name" not in z: z["name"] = f"Zone{i+1}"
            if "enabled" not in z: z["enabled"] = True
            base = BASE_ZONE_DEFS[i]
            try:
                zmin = float(z.get("min", base["fixed_min"]))
                zmax = float(z.get("max", base["fixed_max"]))
            except:
                zmin, zmax = base["fixed_min"], base["fixed_max"]
            if zmin >= zmax:
                zmin, zmax = base["fixed_min"], base["fixed_max"]
            z["min"] = zmin
            z["max"] = zmax

    # Cronus apps launcher config
    cronus_apps = data.get("cronus_apps")
    if not isinstance(cronus_apps, list):
        cronus_apps = []
    # Each entry should be dict {name, path}
    cleaned_apps = []
    for app in cronus_apps:
        if isinstance(app, dict) and "path" in app:
            nm = app.get("name") or os.path.basename(app["path"])
            cleaned_apps.append({"name": nm, "path": app["path"]})
    cronus_apps = cleaned_apps
    cronus_default = data.get("cronus_app_default", 0)
    if not (0 <= cronus_default < len(cronus_apps)):
        cronus_default = 0 if cronus_apps else -1

    return {
        "log_dir": LOG_BASE,
        "ch1": merge_channel(data.get("ch1", {})),
        "ch2": merge_channel(data.get("ch2", {})),
        "show_zones": data.get("show_zones", False),
        "zones": zones_cfg,
        "cronus_apps": cronus_apps,
        "cronus_app_default": cronus_default
    }

def save_config(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception as e:
        print(f"Failed to save config: {e}")

class StatusWorker(QThread):
    status_update = pyqtSignal(bool, str)
    def __init__(self):
        super().__init__()
        self.running = True
    def run(self):
        while self.running:
            try:
                status_response = safe_get_json(f"{API_BASE}/Status")
                connected = status_response is not None and status_response.get("OK", False)
                mode = "Unknown"
                if connected:
                    mode_resp = safe_get_json(f"{API_BASE}/Mode")
                    if mode_resp and mode_resp.get("OK"):
                        mode = mode_resp.get("Mode", "Unknown")
                self.status_update.emit(connected, mode)
            except:
                self.status_update.emit(False, "Unknown")
            time.sleep(3)
    def stop(self):
        self.running = False

class TestWorker(QThread):
    update_status = pyqtSignal(bool, float, float)
    finished = pyqtSignal()
    progress_update = pyqtSignal(int, int)
    power_curve_finished = pyqtSignal(list)
    connection_lost = pyqtSignal()
    command_sent = pyqtSignal(str)
    current_wavelength = pyqtSignal(float)

    def __init__(self, channel: int, params: dict, device_range: tuple):
        super().__init__()
        self.channel = channel
        self.running = False
        self.wait_time = params["wait_time"]
        self.range_min = params["test_min"]
        self.range_max = params["test_max"]
        self.cycles = None if (params["cycles"] is None or params["cycles"] <= 0) else params["cycles"]
        self.measure_power_curve = params["measure_power_curve"]
        self.fail_log_file = None
        self.test_results = []
        self.power_curve_data = []
        self.device_min, self.device_max = device_range

        if self.device_min is None or self.device_max is None:
            self.range_min = None
            self.range_max = None
        else:
            if self.range_min is None or self.range_max is None:
                self.range_min, self.range_max = self.device_min, self.device_max
            self.range_min = max(self.range_min, self.device_min)
            self.range_max = min(self.range_max, self.device_max)
            if self.range_min >= self.range_max:
                self.range_min, self.range_max = self.device_min, self.device_max

    def configure_logs(self):
        ensure_log_dirs()
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.fail_log_file = os.path.join(LOG_BASE, f"Ch{self.channel}_wavelength_failures_{ts}.txt")
        with open(self.fail_log_file, "w") as f:
            f.write(f"Cronus Ch{self.channel} Wavelength Failures Log\n")
            f.write(f"Test started: {datetime.now():%Y-%m-%d %H:%M:%S}\n")
            f.write("-" * 50 + "\n")

    def run(self):
        self.configure_logs()
        self.running = True
        attempts = 0
        self.test_results = []
        if self.range_min is None or self.range_max is None:
            self.finished.emit()
            return
        while self.running and (self.cycles is None or attempts < self.cycles):
            if not check_connection(self.channel):
                self.connection_lost.emit()
                time.sleep(2)
                continue
            wl = round(random.uniform(self.range_min, self.range_max), 1)
            wl = min(max(wl, self.device_min), self.device_max)
            self.current_wavelength.emit(wl)
            success, duration = self._perform_wavelength_attempt(wl)
            attempts += 1
            self.test_results.append({
                "timestamp": datetime.now(),
                "wavelength": wl,
                "wl_success": success,
                "wl_duration": duration
            })
            self.update_status.emit(success, duration, wl)
            self.progress_update.emit(attempts, self.cycles or 0)
            if not success:
                with open(self.fail_log_file, "a") as f:
                    f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} - Failed wavelength: "
                            f"{wl} nm (Duration: {duration:.1f}s)\n")
            time.sleep(self.wait_time + 3)  # post-set wait + margin
        if self.measure_power_curve and self.running and self.range_min is not None:
            self._measure_power_curve()
        self.finished.emit()

    def _perform_wavelength_attempt(self, wl):
        start_time = time.time()
        self.command_sent.emit(f"PUT /Ch{self.channel}/Wavelength: {wl} nm")
        put_resp = safe_put_json(f"{API_BASE}/Ch{self.channel}/Wavelength", {"OK": True, "Wavelength": wl})
        if put_resp is None:
            self.connection_lost.emit()
            return False, time.time() - start_time
        active_started = False
        for _ in range(10):
            if not self.running: break
            self.command_sent.emit(f"GET /Ch{self.channel}/Status")
            st = safe_get_json(f"{API_BASE}/Ch{self.channel}/Status")
            if st is None:
                self.connection_lost.emit()
                time.sleep(0.5); continue
            if st.get("IsWavelengthSettingActive", False):
                active_started = True
                break
            time.sleep(0.5)
        if not active_started:
            return False, time.time() - start_time
        success = False
        for _ in range(120):
            if not self.running: break
            self.command_sent.emit(f"GET /Ch{self.channel}/Status")
            st = safe_get_json(f"{API_BASE}/Ch{self.channel}/Status")
            if st is None:
                self.connection_lost.emit()
                time.sleep(1); continue
            if not st.get("IsWavelengthSettingActive", True):
                if st.get("WavelengthSettingState", "") == "Success":
                    success = True
                break
            time.sleep(1)
        return success, time.time() - start_time

    def _measure_power_curve(self):
        if self.range_min is None or self.range_max is None:
            self.power_curve_finished.emit([])
            return
        self.power_curve_data = []
        span = self.range_max - self.range_min
        if span < 10:
            wls = [int((self.range_min + self.range_max) / 2)]
        else:
            wls = list(range(int(self.range_min), int(self.range_max) + 1, 10))
            if wls[-1] < int(self.range_max):
                wls.append(int(self.range_max))
        for wl in wls:
            if not self.running: break
            if not check_connection(self.channel):
                self.connection_lost.emit()
                time.sleep(2)
                if not check_connection(self.channel): break
            wl = min(max(wl, self.device_min), self.device_max)
            success, _ = self._perform_wavelength_attempt(wl)
            if success:
                time.sleep(3)
                self.command_sent.emit(f"GET /Ch{self.channel}/Power")
                p = safe_get_json(f"{API_BASE}/Ch{self.channel}/Power")
                if p and p.get("OK"):
                    self.power_curve_data.append({"wavelength": wl, "power": p.get("Power", 0.0)})
        self.power_curve_finished.emit(self.power_curve_data)

class SmoothProgressBar(QProgressBar):
    SHOW_PERCENT_TEXT = False
    ANIM_DURATION_MS = 520
    BG_COLOR = QColor("#f1f5f9")
    TRACK_BORDER = QColor("#d8dee5")
    FILL_GRADIENT_START = QColor("#4d7cff")
    FILL_GRADIENT_END = QColor("#8fb0ff")
    GLOSS_GRADIENT_TOP = QColor(255, 255, 255, 110)
    GLOSS_GRADIENT_BOTTOM = QColor(255, 255, 255, 0)
    TEXT_COLOR = QColor("#0f172a")
    def __init__(self,*args,**kwargs):
        super().__init__(*args,**kwargs)
        self._animValue=0.0; self._anim=None
        self.setTextVisible(False); self.setFixedHeight(40)
        self.setAttribute(Qt.WidgetAttribute.WA_OpaquePaintEvent,False)
        self.setAttribute(Qt.WidgetAttribute.WA_NoSystemBackground,True)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground,True)
    def sizeHint(self):
        s=super().sizeHint(); s.setHeight(40); return s
    def setValue(self,v:int):
        v=int(v); old=super().value(); super().setValue(v)
        target=0 if self.maximum()==0 else (v/self.maximum()) if self.maximum()>0 else 0
        if old==0 and v==0:
            self._animValue=0; self.update(); return
        self._animate_to(target)
    def _animate_to(self,target_ratio:float):
        if self._anim and self._anim.state()==QPropertyAnimation.State.Running:
            self._anim.stop()
        self._anim=QPropertyAnimation(self,b"animValue")
        self._anim.setStartValue(self._animValue)
        self._anim.setEndValue(target_ratio)
        self._anim.setDuration(self.ANIM_DURATION_MS)
        self._anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        self._anim.start()
    def getAnimValue(self): return self._animValue
    def setAnimValue(self,val): self._animValue=val; self.update()
    animValue=pyqtProperty(float,fget=getAnimValue,fset=setAnimValue)
    def paintEvent(self,event):
        painter=QPainter(self); painter.setRenderHint(QPainter.RenderHint.Antialiasing,True)
        rect=self.rect().adjusted(1,1,-1,-1); radius=rect.height()/2.0
        painter.setPen(QPen(self.TRACK_BORDER,1)); painter.setBrush(QBrush(self.BG_COLOR))
        painter.drawRoundedRect(rect,radius,radius)
        pct=self._animValue if self.maximum()>0 else 0
        fill_w=max(0.0,rect.width()*pct)
        if fill_w>2:
            fill_rect=QRectF(rect.x(),rect.y(),fill_w,rect.height())
            grad=QLinearGradient(fill_rect.topLeft(),fill_rect.topRight())
            grad.setColorAt(0,self.FILL_GRADIENT_START); grad.setColorAt(1,self.FILL_GRADIENT_END)
            painter.setPen(Qt.PenStyle.NoPen); painter.setBrush(QBrush(grad))
            painter.drawRoundedRect(fill_rect,radius,radius)
            gloss=QLinearGradient(fill_rect.topLeft(),fill_rect.bottomLeft())
            gloss.setColorAt(0,self.GLOSS_GRADIENT_TOP); gloss.setColorAt(1,self.GLOSS_GRADIENT_BOTTOM)
            painter.setBrush(QBrush(gloss)); painter.drawRoundedRect(fill_rect,radius,radius)
        if self.SHOW_PERCENT_TEXT and self.maximum()>0:
            percent=(self.value()/self.maximum())*100
            painter.setPen(self.TEXT_COLOR); font=painter.font()
            font.setPointSize(11); font.setBold(True); painter.setFont(font)
            painter.drawText(rect,Qt.AlignmentFlag.AlignCenter,f"{percent:.0f}%")
        painter.end()

class ChannelPanel(QFrame):
    STATUS_COLORS = {
        "Standby": "#27ae60",
        "Running": "#34495e",
        "Power Curve": "#eab308",
        "Completed": "#27ae60",
        "Range Unknown": "#f59e0b",
        "Error": "#e74c3c"
    }
    def __init__(self,channel:int,get_params_callable,get_device_range_callable):
        super().__init__()
        self.channel=channel
        self.get_params_callable=get_params_callable
        self.get_device_range_callable=get_device_range_callable
        self.worker=None; self.aborted=False; self.needs_reset_next_start=False
        self.timer=QTimer(); self.timer.timeout.connect(self._update_timer)
        self.elapsed_seconds=0
        self.success_count=0; self.fail_count=0; self.total_time=0.0
        self.durations=[]; self.wavelengths=[]; self.success_wavelengths=[]; self.failed_wavelengths=[]
        self.power_curve_data=[]; self.command_log=[]
        self.is_test_completed=False; self.wavelength_tested=None
        self.total_cycles=None; self.current_wait_time=0.0; self.latest_eta="-"
        self.device_min,self.device_max=self.get_device_range_callable(self.channel)
        self._apply_style(); self._build_ui()
        self.refresh_param_summary()
        self._set_status("Range Unknown" if self.device_min is None else "Standby")
    def _apply_style(self):
        self.setStyleSheet("""
            QFrame {background:#ffffff;border:1px solid #e3e8ef;border-radius:24px;margin:10px;}
            QToolTip {background:#1e293b;color:#f1f5f9;border:1px solid #334155;padding:4px;
                      border-radius:6px;font-size:11px;}
        """)
    def _build_ui(self):
        main=QVBoxLayout(self); main.setSpacing(14); main.setContentsMargins(14,14,14,14)
        header=QFrame(); header.setStyleSheet("""
            QFrame {border-radius:18px;background:qlineargradient(x1:0,y1:0,x2:1,y2:0,
            stop:0 #eef2f7, stop:1 #f8fafc);border:1px solid #e2e8f0;}
        """)
        h_l=QHBoxLayout(header); h_l.setContentsMargins(14,8,14,8)
        title_box=QVBoxLayout()
        self.title_label=QLabel(f"Channel {self.channel}")
        self.title_label.setFont(QFont("Segoe UI",18,QFont.Weight.Bold))
        self.attempts_label=QLabel("Attempts: 0")
        self.attempts_label.setStyleSheet("color:#64748b;font-size:11px;font-weight:600;")
        title_box.addWidget(self.title_label); title_box.addWidget(self.attempts_label)
        h_l.addLayout(title_box); h_l.addStretch()
        self.status_pill=QLabel("Standby"); self.status_pill.setStyleSheet(self._status_pill_style("#27ae60"))
        h_l.addWidget(self.status_pill); main.addWidget(header)
        self.param_summary=QLabel()
        self.param_summary.setStyleSheet("color:#475569;font-size:11px;background:#ffffff;padding:8px 10px;"
                                         "border:1px solid #e2e8f0;border-radius:12px;line-height:150%;")
        self.param_summary.setWordWrap(True); main.addWidget(self.param_summary)
        self.progress_bar=SmoothProgressBar(); main.addWidget(self.progress_bar)
        self.progress_info_label=QLabel(""); self.progress_info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progress_info_label.setStyleSheet("font-size:12px;font-weight:600;color:#1e293b;")
        self.progress_info_label.setVisible(False); main.addWidget(self.progress_info_label)
        btn_row=QHBoxLayout()
        self.start_btn=QPushButton("Start"); self.stop_btn=QPushButton("Stop"); self.reset_btn=QPushButton("Reset")
        self.logs_btn=QPushButton("Logs"); self.report_btn=QPushButton("Report"); self.export_btn=QPushButton("Export")
        for b in [self.start_btn,self.stop_btn,self.reset_btn,self.logs_btn,self.report_btn,self.export_btn]:
            b.setStyleSheet(self._button_style()); btn_row.addWidget(b)
        self.report_btn.setEnabled(False); self.export_btn.setEnabled(False); main.addLayout(btn_row)
        self.time_combo_label=QLabel("Elapsed: 0s   ETA: -")
        self.time_combo_label.setStyleSheet("font-weight:600;color:#1e293b;font-size:12px;")
        main.addWidget(self.time_combo_label)
        stats_chart=QHBoxLayout(); stats_col=QVBoxLayout()
        self.success_label=self._chip_label("Success: 0","#e8f8ef","#1f5132")
        self.fail_label=self._chip_label("Failed: 0","#fdecec","#7f1d1d")
        self.avg_label=self._chip_label("Avg: 0.0s","#f1f5f9","#475569")
        self.rate_label=self._chip_label("Rate: 0%","#f1f5f9","#475569")
        drop_row1=QHBoxLayout()
        self.success_dropdown_btn=QPushButton("▼"); self.success_dropdown_btn.setFixedWidth(26)
        self.success_dropdown_btn.setStyleSheet(self._dropdown_button_style())
        self.success_dropdown_btn.clicked.connect(self._show_success_wavelengths)
        drop_row1.addWidget(self.success_label); drop_row1.addWidget(self.success_dropdown_btn); drop_row1.addStretch()
        drop_row2=QHBoxLayout()
        self.fail_dropdown_btn=QPushButton("▼"); self.fail_dropdown_btn.setFixedWidth(26)
        self.fail_dropdown_btn.setStyleSheet(self._dropdown_button_style())
        self.fail_dropdown_btn.clicked.connect(self._show_fail_wavelengths)
        drop_row2.addWidget(self.fail_label); drop_row2.addWidget(self.fail_dropdown_btn); drop_row2.addStretch()
        stats_col.addLayout(drop_row1); stats_col.addLayout(drop_row2)
        stats_col.addWidget(self.avg_label); stats_col.addWidget(self.rate_label); stats_col.addStretch()
        stats_chart.addLayout(stats_col)
        self.fig,self.ax=plt.subplots(figsize=(3.4,2.7))
        self.fig.patch.set_facecolor('#ffffff'); self.canvas=FigureCanvas(self.fig)
        stats_chart.addWidget(self.canvas); main.addLayout(stats_chart)
        cmd_label=QLabel("Commands (Last 5)")
        cmd_label.setStyleSheet("font-weight:600;font-size:11px;margin-top:4px;")
        main.addWidget(cmd_label)
        self.command_log_display=QTextEdit(); self.command_log_display.setReadOnly(True)
        self.command_log_display.setMaximumHeight(100)
        self.command_log_display.setStyleSheet("""
            QTextEdit {background:#1e293b;color:#e2e8f0;border:1px solid #334155;border-radius:12px;
                       font-family:Consolas,monospace;font-size:10px;padding:6px;}
        """)
        main.addWidget(self.command_log_display)
        self._update_chart()
        self.start_btn.clicked.connect(self.on_start)
        self.stop_btn.clicked.connect(self.on_stop)
        self.reset_btn.clicked.connect(self.on_reset)
        self.logs_btn.clicked.connect(self.on_open_logs)
        self.report_btn.clicked.connect(self.on_generate_report)
        self.export_btn.clicked.connect(self.on_export_data)
        self.progress_bar.setVisible(False)
    def _chip_label(self,text,bg,fg):
        lab=QLabel(text); lab.setStyleSheet(
            f"QLabel {{ background:{bg}; color:{fg}; padding:4px 10px; border-radius:14px;"
            " font-size:11px; font-weight:600;}}"); return lab
    def _status_pill_style(self,color):
        return f"QLabel {{ background:{color}; color:white; padding:6px 16px; border-radius:18px;" \
               f" font-size:12px; font-weight:600; letter-spacing:0.4px; }}"
    def _set_status(self,text):
        color=self.STATUS_COLORS.get(text,"#34495e")
        self.status_pill.setText(text); self.status_pill.setStyleSheet(self._status_pill_style(color))
    def _button_style(self):
        return """
            QPushButton {background:#34495e;color:white;border:none;padding:8px 16px;border-radius:12px;
                         font-weight:600;font-size:12px;}
            QPushButton:hover {background:#4d6070;}
            QPushButton:disabled {background:#b4c0cb;color:#e2e8f0;}
        """
    def _dropdown_button_style(self):
        return """
            QPushButton {background:#e2e8f0;color:#1e293b;border:none;border-radius:6px;font-weight:600;
                        font-size:10px;padding:2px 4px;}
            QPushButton:hover {background:#cbd5e1;}
        """
    def refresh_param_summary(self):
        params=self.get_params_callable(self.channel)
        self.device_min,self.device_max=self.get_device_range_callable(self.channel)
        if self.device_min is None or self.device_max is None:
            range_text="Range: -"; dev_text="Device Range: -"; self.start_btn.setEnabled(False)
        else:
            if params['test_min'] is None or params['test_max'] is None:
                range_text="Range: -"
            else:
                range_text=f"Range: {params['test_min']:.1f}–{params['test_max']:.1f} nm"
            dev_text=f"Device Range: {self.device_min:.1f}–{self.device_max:.1f} nm"
            if not (self.worker and self.worker.isRunning()):
                self.start_btn.setEnabled(True)
        summary=(f"{range_text}  |  Post-set Wait {params['wait_time']:.1f}s  |  "
                 f"Cycles {params['cycles'] if params['cycles'] and params['cycles']>0 else '∞'}  |  "
                 f"PowerCurve {'Yes' if params['measure_power_curve'] else 'No'}\n{dev_text}")
        self.param_summary.setText(summary)
    def _update_timer(self):
        self.elapsed_seconds+=1
        h=self.elapsed_seconds//3600; m=(self.elapsed_seconds%3600)//60; s=self.elapsed_seconds%60
        elapsed_str=f"{h:02d}:{m:02d}:{s:02d}" if h else f"{m:02d}:{s:02d}"
        self.time_combo_label.setText(f"Elapsed: {elapsed_str}   ETA: {self.latest_eta}")
    def _update_eta(self):
        if not self.total_cycles: self.latest_eta="∞"; return
        attempts_done=self.success_count+self.fail_count
        remaining=self.total_cycles-attempts_done
        if remaining<=0: self.latest_eta="0s"; return
        if not self.durations: self.latest_eta="-"; return
        avg_set_time=sum(self.durations)/len(self.durations)
        per_cycle_est=avg_set_time+self.current_wait_time+3
        rem=remaining*per_cycle_est
        if rem<3600:
            m=int(rem//60); s=int(rem%60); self.latest_eta=f"{m:02d}:{s:02d}"
        else:
            h=int(rem//3600); m=int((rem%3600)//60); s=int(rem%60)
            self.latest_eta=f"{h:02d}:{m:02d}:{s:02d}"
    def _update_chart(self):
        self.ax.clear(); total=self.success_count+self.fail_count
        self.ax.set_aspect('equal'); ring_width=0.23
        if total>0:
            sizes=[self.success_count,self.fail_count]; colors_ring=['#34C759','#FF3B30']
            self.ax.pie(sizes,startangle=90,colors=colors_ring,radius=1.0,counterclock=False,
                        labels=None,wedgeprops=dict(width=ring_width,edgecolor='#ffffff',linewidth=2,antialiased=True))
            shadow=plt.Circle((0,0),1.0,color='black',alpha=0.04,zorder=0); self.ax.add_patch(shadow)
            success_rate=(self.success_count/total)*100
            self.ax.text(0.5,0.5,f"{success_rate:.0f}%",ha='center',va='center',
                         fontsize=26,fontweight='bold',color='#0f172a')
        else:
            placeholder=plt.Circle((0,0),1,color='#eef2f7',zorder=0); self.ax.add_patch(placeholder)
            self.ax.text(0.5,0.5,"–%",ha='center',va='center',fontsize=24,fontweight='bold',color='#94a3b8')
        self.ax.set_xlim(-1.1,1.1); self.ax.set_ylim(-1.1,1.1); self.ax.axis('off'); self.canvas.draw()
    def _clear_statistics(self):
        self.success_count=0; self.fail_count=0; self.total_time=0.0; self.durations.clear()
        self.wavelengths.clear(); self.success_wavelengths.clear(); self.failed_wavelengths.clear()
        self.power_curve_data.clear(); self.command_log.clear(); self.command_log_display.clear()
        self.is_test_completed=False; self.attempts_label.setText("Attempts: 0")
        self.avg_label.setText("Avg: 0.0s"); self.success_label.setText("Success: 0")
        self.fail_label.setText("Failed: 0"); self.rate_label.setText("Rate: 0%")
        self.elapsed_seconds=0; self.latest_eta="-"
        self.time_combo_label.setText("Elapsed: 0s   ETA: -")
        self._update_chart(); self.report_btn.setEnabled(False); self.export_btn.setEnabled(False)
    def _show_success_wavelengths(self):
        if not self.success_wavelengths: return
        menu=QMenu(self)
        for wl in list(self.success_wavelengths)[-150:][::-1]:
            pass  # unreachable; corrected below
        for wl in list(self.success_wavelengths)[-150:][::-1]:
            menu.addAction(f"{wl:.1f} nm")
        menu.exec(self.success_dropdown_btn.mapToGlobal(self.success_dropdown_btn.rect().bottomLeft()))
    def _show_fail_wavelengths(self):
        if not self.failed_wavelengths: return
        menu=QMenu(self)
        for wl in list(self.failed_wavelengths)[-150:][::-1]:
            menu.addAction(f"{wl:.1f} nm")
        menu.exec(self.fail_dropdown_btn.mapToGlobal(self.fail_dropdown_btn.rect().bottomLeft()))
    def on_start(self):
        if self.worker and self.worker.isRunning(): return
        params=self.get_params_callable(self.channel)
        device_min,device_max=self.get_device_range_callable(self.channel)
        if device_min is None or device_max is None:
            QMessageBox.warning(self,"Device Range Unknown",
                                "Cannot start: range unknown. Click Reconnect or read in Settings.")
            self._set_status("Range Unknown"); return
        if params['test_min'] is None or params['test_max'] is None:
            params['test_min'],params['test_max']=device_min,device_max
        else:
            params['test_min']=max(params['test_min'],device_min)
            params['test_max']=min(params['test_max'],device_max)
            if params['test_min']>=params['test_max']:
                params['test_min'],params['test_max']=device_min,device_max
        if self.needs_reset_next_start or (self.success_count or self.fail_count):
            self._clear_statistics()
        self.needs_reset_next_start=False; self.aborted=False
        self.refresh_param_summary()
        self.worker=TestWorker(self.channel,params,(device_min,device_max))
        self.total_cycles=params['cycles'] if params['cycles'] and params['cycles']>0 else None
        self.current_wait_time=params['wait_time']
        if self.total_cycles:
            self.progress_bar.setMaximum(self.total_cycles); self.progress_bar.setValue(0)
            self.progress_bar.setVisible(True); self.progress_info_label.setVisible(True)
            self.progress_info_label.setText(f"0 / {self.total_cycles} (0%)")
        else:
            self.progress_bar.setVisible(False); self.progress_info_label.setVisible(False)
        self.worker.update_status.connect(self.on_result)
        self.worker.progress_update.connect(self.on_progress)
        self.worker.finished.connect(self.on_worker_finished)
        self.worker.power_curve_finished.connect(self.on_power_curve_finished)
        self.worker.command_sent.connect(self.on_command_sent)
        self.worker.start(); self._set_status("Running")
        self.elapsed_seconds=0; self.latest_eta="-" if self.total_cycles else "∞"
        self.timer.start(1000); self.start_btn.setEnabled(False)
    def on_stop(self):
        if not self.worker or not self.worker.isRunning(): return
        self.aborted=True; self.worker.running=False; self.worker.wait(150)
        if self.timer.isActive(): self.timer.stop()
        self._set_status("Standby" if self.device_min is not None else "Range Unknown")
        self.start_btn.setEnabled(True); self.latest_eta="-"
        self.time_combo_label.setText("Elapsed: 0s   ETA: -"); self.needs_reset_next_start=True
    def on_reset(self):
        if self.worker and self.worker.isRunning():
            self.aborted=True; self.worker.running=False; self.worker.wait(150)
        if self.timer.isActive(): self.timer.stop()
        self._clear_statistics(); self.needs_reset_next_start=False
        self._set_status("Standby" if self.device_min is not None else "Range Unknown")
        self.start_btn.setEnabled(self.device_min is not None)
        self.progress_bar.setVisible(False); self.progress_info_label.setVisible(False)
    def on_command_sent(self,cmd):
        ts=datetime.now().strftime("%H:%M:%S"); entry=f"[{ts}] {cmd}"
        if self.command_log and self.command_log[0]==entry: return
        self.command_log.insert(0,entry)
        if len(self.command_log)>5: self.command_log.pop()
        self.command_log_display.setText("\n".join(self.command_log))
    def on_result(self,success,duration,wavelength):
        if self.aborted: return
        if success:
            self.success_count+=1; self.total_time+=duration; self.success_wavelengths.append(wavelength)
        else:
            self.fail_count+=1; self.failed_wavelengths.append(wavelength)
        self.durations.append(duration); self.wavelengths.append(wavelength)
        attempts_done=self.success_count+self.fail_count
        self.attempts_label.setText(f"Attempts: {attempts_done}")
        if self.wavelength_tested: self.wavelength_tested(wavelength,success)
        self.success_label.setText(f"Success: {self.success_count}")
        self.fail_label.setText(f"Failed: {self.fail_count}")
        avg=self.total_time/self.success_count if self.success_count else 0.0
        self.avg_label.setText(f"Avg: {avg:.1f}s")
        rate=(self.success_count/attempts_done*100) if attempts_done else 0
        self.rate_label.setText(f"Rate: {rate:.1f}%")
        self._update_chart(); self._update_eta(); self._update_timer()
    def on_progress(self,current,total):
        if self.aborted: return
        if total>0:
            self.progress_bar.setValue(current)
            percent=current/total*100
            self.progress_info_label.setText(f"{current} / {total} ({percent:.0f}%)")
    def on_worker_finished(self):
        if self.aborted: return
        params=self.get_params_callable(self.channel)
        if params['measure_power_curve']:
            self._set_status("Power Curve")
        else:
            self._finalize_complete()
    def on_power_curve_finished(self,data):
        if self.aborted: return
        self.power_curve_data=data; self._finalize_complete()
    def _finalize_complete(self):
        if self.aborted: return
        self._set_status("Completed")
        if self.timer.isActive(): self.timer.stop()
        self.progress_bar.setVisible(False); self.progress_info_label.setVisible(False)
        self.is_test_completed=True; self.report_btn.setEnabled(True); self.export_btn.setEnabled(True)
        self.start_btn.setEnabled(True); self.latest_eta="0s"; self._update_timer()
        self.needs_reset_next_start=True
    def on_open_logs(self):
        if os.path.exists(LOG_BASE):
            try: os.startfile(LOG_BASE)
            except Exception: pass
    def on_export_data(self):
        if not (self.worker and self.worker.test_results): return
        ts=datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name=f"Ch{self.channel}_TestResults_{ts}.csv"
        filename,_=QFileDialog.getSaveFileName(self,"Export Test Data",
                                               os.path.join(LOG_BASE,default_name),
                                               "CSV Files (*.csv);;All Files (*)")
        if not filename: return
        try:
            with open(filename,"w",newline='',encoding="utf-8") as f:
                w=csv.writer(f)
                w.writerow(["Timestamp","Wavelength (nm)","Success","Duration (s)"])
                for r in self.worker.test_results:
                    w.writerow([
                        r["timestamp"].strftime("%Y-%m-%d %H:%M:%S"),
                        f"{r['wavelength']:.1f}",
                        "Yes" if r['wl_success'] else "No",
                        f"{r['wl_duration']:.2f}"
                    ])
            self._set_status("Completed")
        except Exception:
            self._set_status("Error")
    def on_generate_report(self):
        if not self.is_test_completed or not (self.worker and self.worker.test_results): return
        ts=datetime.now().strftime("%Y%m%d_%H%M%S")
        filename=os.path.join(LOG_BASE,f"Ch{self.channel}_Report_{ts}.pdf")
        self._create_pdf_report(filename)
    def _create_pdf_report(self,filename):
        doc=SimpleDocTemplate(filename,pagesize=letter)
        styles=getSampleStyleSheet(); story=[]
        title_style=ParagraphStyle('Title',parent=styles['Heading1'],fontSize=20,
                                   textColor=colors.HexColor('#2c3e50'),alignment=1,
                                   spaceAfter=25,fontName='Helvetica-Bold')
        story.append(Paragraph(f"Cronus App - Channel {self.channel} Test Report",title_style))
        story.append(Spacer(1,12))
        total_tests=self.success_count+self.fail_count
        summary_data=[
            ["Parameter","Value"],
            ["Total Wavelength Tests",str(total_tests)],
            ["Successful",str(self.success_count)],
            ["Failed",str(self.fail_count)],
            ["Success Rate",f"{(self.success_count/total_tests*100):.1f}%" if total_tests else "0%"],
            ["Average Duration",f"{(self.total_time/self.success_count):.1f}s" if self.success_count else "N/A"],
        ]
        table=Table(summary_data)
        table.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#34495e')),
            ('TEXTCOLOR',(0,0),(-1,0),colors.white),
            ('ALIGN',(0,0),(-1,-1),'CENTER'),
            ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
            ('BACKGROUND',(0,1),(-1,-1),colors.HexColor('#ecf0f1')),
            ('TEXTCOLOR',(0,1),(-1,-1),colors.HexColor('#2c3e50')),
            ('GRID',(0,0),(-1,-1),0.4,colors.HexColor('#bdc3c7'))
        ]))
        story.append(Paragraph("Summary",styles['Heading2']))
        story.append(table); story.append(Spacer(1,16))
        failed=[r for r in self.worker.test_results if not r['wl_success']]
        if failed:
            story.append(Paragraph("Failed Wavelengths",styles['Heading2']))
            failed_rows=[["Timestamp","Wavelength (nm)","Duration (s)"]]
            for r in failed:
                failed_rows.append([
                    r['timestamp'].strftime("%H:%M:%S"),
                    f"{r['wavelength']:.1f}",
                    f"{r['wl_duration']:.1f}"
                ])
            ft=Table(failed_rows)
            ft.setStyle(TableStyle([
                ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#95a5a6')),
                ('TEXTCOLOR',(0,0),(-1,0),colors.white),
                ('ALIGN',(0,0),(-1,-1),'CENTER'),
                ('GRID',(0,0),(-1,-1),0.4,colors.HexColor('#d3dce6'))
            ]))
            story.append(ft); story.append(Spacer(1,12))
        if self.power_curve_data:
            story.append(Paragraph("Power Curve",styles['Heading2']))
            fig=Figure(figsize=(7,4.5)); ax=fig.add_subplot(111)
            wls=[p['wavelength'] for p in self.power_curve_data]
            pows=[p['power'] for p in self.power_curve_data]
            ax.plot(wls,pows,'o-',color='#34495e',linewidth=2,markersize=6,
                    markerfacecolor='#95a5a6',markeredgecolor='#2c3e50',markeredgewidth=1.2)
            ax.set_xlabel("Wavelength (nm)"); ax.set_ylabel("Power (mW)")
            ax.set_title("Power vs Wavelength"); ax.grid(True,alpha=0.25); fig.tight_layout()
            buf=io.BytesIO(); fig.savefig(buf,format='png',dpi=130,bbox_inches='tight'); buf.seek(0)
            story.append(Image(buf,width=5*inch,height=3.5*inch)); story.append(Spacer(1,12))
        doc.build(story)
        try: subprocess.Popen(f'start "" "{filename}"',shell=True)
        except Exception: pass

class SettingsDialog(QDialog):
    def __init__(self,parent,show_map,show_zones,log_dir,ch1_params,ch2_params,dev_ranges,zones_cfg,cronus_apps,cronus_default_index):
        super().__init__(parent)
        self.setWindowTitle("Settings"); self.setModal(True); self.setMinimumWidth(1100)
        self.show_map=show_map; self.show_zones=show_zones; self.log_dir=log_dir
        self.ch1_params=ch1_params.copy(); self.ch2_params=ch2_params.copy()
        self.dev_ranges={1:dev_ranges.get(1,(None,None)),2:dev_ranges.get(2,(None,None))}
        self.zones_cfg=[z.copy() for z in zones_cfg]
        self.cronus_apps=[a.copy() for a in cronus_apps]
        self.cronus_default_index=cronus_default_index if 0<=cronus_default_index<len(self.cronus_apps) else -1
        self.app_rows=[]
        self._build_ui()
    def _build_ui(self):
        layout=QVBoxLayout(self); layout.setContentsMargins(28,28,28,28); layout.setSpacing(20)
        title=QLabel("Application Settings"); title.setFont(QFont("Segoe UI",20,QFont.Weight.Bold))
        layout.addWidget(title)
        # Log directory
        lg_group=QGroupBox("Log Directory")
        lg_layout=QHBoxLayout(lg_group)
        self.log_dir_input=QLineEdit(self.log_dir); self.log_dir_input.setReadOnly(True)
        browse_btn=QPushButton("Browse"); browse_btn.clicked.connect(self._browse_dir)
        lg_layout.addWidget(self.log_dir_input); lg_layout.addWidget(browse_btn)
        layout.addWidget(lg_group)
        # Display
        disp_group=QGroupBox("Display Options"); disp_layout=QVBoxLayout(disp_group)
        self.map_checkbox=QCheckBox("Show wavelength testing map"); self.map_checkbox.setChecked(self.show_map)
        self.zones_checkbox=QCheckBox("Display zones in the map"); self.zones_checkbox.setChecked(self.show_zones)
        self._style_checkbox(self.map_checkbox); self._style_checkbox(self.zones_checkbox)
        disp_layout.addWidget(self.map_checkbox); disp_layout.addWidget(self.zones_checkbox)
        layout.addWidget(disp_group)
        # Channels
        self.ch1_group=self._channel_params_group("Channel 1 Parameters",self.ch1_params,1)
        self.ch2_group=self._channel_params_group("Channel 2 Parameters",self.ch2_params,2)
        layout.addWidget(self.ch1_group); layout.addWidget(self.ch2_group)
        # Zones
        zones_group=QGroupBox("Wavelength Map Zones (Name / Min / Max / Enable)")
        zg_layout=QVBoxLayout(zones_group)
        self.zone_rows=[]
        for i,base in enumerate(BASE_ZONE_DEFS):
            zcfg=self.zones_cfg[i] if i<len(self.zones_cfg) else {
                "name":f"Zone{i+1}","enabled":True,"min":base["fixed_min"],"max":base["fixed_max"]
            }
            row=QHBoxLayout()
            label=QLabel(f"Base Default: {base['fixed_min']:.1f} – {base['fixed_max']:.1f} nm")
            label.setStyleSheet("color:#475569;font-size:11px;")
            name_edit=QLineEdit(zcfg.get("name","")); name_edit.setFixedWidth(120)
            min_edit=QLineEdit(f"{float(zcfg.get('min',base['fixed_min'])):.3f}"); min_edit.setFixedWidth(90)
            max_edit=QLineEdit(f"{float(zcfg.get('max',base['fixed_max'])):.3f}"); max_edit.setFixedWidth(90)
            enable_cb=QCheckBox("Enabled"); enable_cb.setChecked(zcfg.get("enabled",True))
            self._style_checkbox(enable_cb)
            def make_default(idx=i):
                b=BASE_ZONE_DEFS[idx]
                self.zone_rows[idx]["min_edit"].setText(f"{b['fixed_min']:.3f}")
                self.zone_rows[idx]["max_edit"].setText(f"{b['fixed_max']:.3f}")
            def_btn=QPushButton("Default Range"); def_btn.setFixedWidth(110)
            def_btn.setStyleSheet("""
                QPushButton {background:#3b82f6;color:#ffffff;border:none;padding:4px 8px;
                             border-radius:6px;font-size:11px;font-weight:600;}
                QPushButton:hover {background:#2563eb;}
            """)
            def_btn.clicked.connect(make_default)
            row.addWidget(label); row.addSpacing(8)
            row.addWidget(QLabel("Name:")); row.addWidget(name_edit)
            row.addSpacing(6); row.addWidget(QLabel("Min:")); row.addWidget(min_edit)
            row.addSpacing(6); row.addWidget(QLabel("Max:")); row.addWidget(max_edit)
            row.addSpacing(6); row.addWidget(enable_cb)
            row.addSpacing(6); row.addWidget(def_btn); row.addStretch()
            zg_layout.addLayout(row)
            self.zone_rows.append({"name_edit":name_edit,"min_edit":min_edit,"max_edit":max_edit,"enable_cb":enable_cb})
        layout.addWidget(zones_group)
        # Cronus App Launcher
        app_group=QGroupBox("Cronus App Launcher")
        app_layout=QVBoxLayout(app_group)
        header_row=QHBoxLayout()
        header_row.addWidget(QLabel("Manage external Cronus application entries (exe or .lnk)."))
        header_row.addStretch()
        add_btn=QPushButton("Add App")
        add_btn.setStyleSheet("""
            QPushButton {background:#3b82f6;color:white;border:none;padding:6px 14px;border-radius:8px;
                         font-weight:600;font-size:12px;}
            QPushButton:hover {background:#2563eb;}
        """)
        add_btn.clicked.connect(self._add_app_entry)
        header_row.addWidget(add_btn)
        app_layout.addLayout(header_row)
        self.apps_container=QVBoxLayout()
        self._rebuild_apps_ui()
        app_layout.addLayout(self.apps_container)
        layout.addWidget(app_group)
        layout.addStretch()
        btns=QHBoxLayout(); btns.addStretch()
        cancel=QPushButton("Cancel"); save=QPushButton("Apply")
        for b in (cancel,save):
            b.setStyleSheet("""
                QPushButton {background:#64748b;color:white;border:none;padding:10px 22px;
                             border-radius:10px;font-weight:600;}
                QPushButton:hover {background:#475569;}
            """)
        cancel.clicked.connect(self.reject); save.clicked.connect(self._on_apply)
        btns.addWidget(cancel); btns.addWidget(save)
        layout.addLayout(btns)
    def _style_checkbox(self,cb:QCheckBox):
        cb.setStyleSheet("""
            QCheckBox {font-size:13px;font-weight:600;color:#334155;padding:4px;}
            QCheckBox::indicator {width:20px;height:20px;border-radius:6px;
                                  border:2px solid #64748b;background:#ffffff;}
            QCheckBox::indicator:checked {background:#2563eb;border:2px solid #1d4ed8;}
        """)
    def _channel_params_group(self,title,params,ch):
        g=QGroupBox(title); g_layout=QVBoxLayout(g)
        dmin,dmax=self.dev_ranges.get(ch,(None,None))
        dev_text="Device Range: -" if dmin is None or dmax is None else f"Device Range: {dmin:.1f} – {dmax:.1f} nm"
        top=QHBoxLayout(); dev_label=QLabel(dev_text); dev_label.setStyleSheet("color:#475569;font-size:11px;")
        read_btn=QPushButton("Read Min/Max"); read_btn.setStyleSheet("""
            QPushButton {background:#3b82f6;color:white;border:none;padding:4px 12px;border-radius:6px;
                         font-size:11px;font-weight:600;}
            QPushButton:hover {background:#2563eb;}
        """)
        read_btn.clicked.connect(lambda _,c=ch,lbl=dev_label: self._read_range(c,lbl))
        top.addWidget(dev_label); top.addStretch(); top.addWidget(read_btn); g_layout.addLayout(top)
        row1=QHBoxLayout()
        self_min=QLineEdit("" if params["test_min"] is None else f"{params['test_min']}"); self_min.setFixedWidth(90)
        self_max=QLineEdit("" if params["test_max"] is None else f"{params['test_max']}"); self_max.setFixedWidth(90)
        row1.addWidget(QLabel("Min:")); row1.addWidget(self_min)
        row1.addWidget(QLabel("Max:")); row1.addWidget(self_max)
        row2=QHBoxLayout()
        wait_in=QLineEdit(f"{params['wait_time']}"); wait_in.setFixedWidth(80)
        cyc_in=QSpinBox(); cyc_in.setRange(0,10000000)
        cyc_in.setValue(params['cycles'] if params['cycles'] is not None else 0)
        row2.addWidget(QLabel("Post-set Wait (s):")); row2.addWidget(wait_in)
        row2.addWidget(QLabel("Cycles (0=∞):")); row2.addWidget(cyc_in)
        pc_box=QCheckBox("Measure power curve after test"); pc_box.setChecked(params["measure_power_curve"])
        self._style_checkbox(pc_box)
        g._dev_label=dev_label; g._min_input=self_min; g._max_input=self_max
        g._wait_input=wait_in; g._cycles_input=cyc_in; g._pc_box=pc_box; g._channel=ch
        g_layout.addLayout(row1); g_layout.addLayout(row2); g_layout.addWidget(pc_box)
        return g
    def _read_range(self,channel,dev_label:QLabel):
        dmin,dmax=fetch_device_range(channel)
        if dmin is None or dmax is None:
            QMessageBox.warning(self,"Read Failed",f"Could not read device range for Channel {channel}."); return
        self.dev_ranges[channel]=(dmin,dmax)
        dev_label.setText(f"Device Range: {dmin:.1f} – {dmax:.1f} nm")
        grp=self.ch1_group if channel==1 else self.ch2_group
        try: cur_min=float(grp._min_input.text()) if grp._min_input.text().strip() else None
        except: cur_min=None
        try: cur_max=float(grp._max_input.text()) if grp._max_input.text().strip() else None
        except: cur_max=None
        if cur_min is None or cur_max is None or cur_min>=cur_max:
            grp._min_input.setText(f"{dmin}"); grp._max_input.setText(f"{dmax}")
        QMessageBox.information(self,"Range Updated",f"Channel {channel} range set to {dmin:.1f} – {dmax:.1f} nm")
    def _browse_dir(self):
        directory=QFileDialog.getExistingDirectory(self,"Select Log Directory",self.log_dir)
        if directory: self.log_dir_input.setText(directory)
    def _sanitize_channel(self,group:QGroupBox,current_params:dict):
        ch=group._channel; dmin,dmax=self.dev_ranges.get(ch,(None,None)); messages=[]
        txt_min=group._min_input.text().strip(); txt_max=group._max_input.text().strip()
        if txt_min=="" or dmin is None: user_min=None
        else:
            try: user_min=float(txt_min)
            except: user_min=None; messages.append(f"Channel {ch}: Invalid min -> cleared")
        if txt_max=="" or dmax is None: user_max=None
        else:
            try: user_max=float(txt_max)
            except: user_max=None; messages.append(f"Channel {ch}: Invalid max -> cleared")
        if dmin is not None and dmax is not None:
            if user_min is None or user_min<dmin: user_min=dmin
            if user_max is None or user_max>dmax: user_max=dmax
            if user_min>=user_max: user_min,user_max=dmin,dmax
        try:
            wait_time=float(group._wait_input.text())
            if wait_time<0: messages.append(f"Channel {ch}: Negative post-set wait -> 0"); wait_time=0.0
        except:
            wait_time=current_params.get("wait_time",3.0); messages.append(f"Channel {ch}: Invalid post-set wait -> {wait_time}")
        cycles=group._cycles_input.value(); power_curve=group._pc_box.isChecked()
        return {
            "test_min": user_min,
            "test_max": user_max,
            "wait_time": wait_time,
            "cycles": cycles,
            "measure_power_curve": power_curve
        }, messages
    def _on_apply(self):
        self.show_map=self.map_checkbox.isChecked()
        self.show_zones=self.zones_checkbox.isChecked()
        self.log_dir=self.log_dir_input.text().strip() or self.log_dir
        ch1_new,msg1=self._sanitize_channel(self.ch1_group,self.ch1_params)
        ch2_new,msg2=self._sanitize_channel(self.ch2_group,self.ch2_params)
        self.ch1_params=ch1_new; self.ch2_params=ch2_new
        new_zones=[]
        for i,row in enumerate(self.zone_rows):
            base=BASE_ZONE_DEFS[i]
            name=row["name_edit"].text().strip() or f"Zone{i+1}"
            try: zmin=float(row["min_edit"].text().strip())
            except: zmin=base["fixed_min"]
            try: zmax=float(row["max_edit"].text().strip())
            except: zmax=base["fixed_max"]
            if zmin>=zmax: zmin,zmax=base["fixed_min"],base["fixed_max"]
            enabled=row["enable_cb"].isChecked()
            new_zones.append({"name":name,"enabled":enabled,"min":zmin,"max":zmax})
        self.zones_cfg=new_zones
        # Apps
        new_apps=[]
        default_index=-1
        for idx,row in enumerate(self.app_rows):
            name=row["name_edit"].text().strip() or os.path.basename(row["path_edit"].text().strip() or f"App{idx+1}")
            path=row["path_edit"].text().strip()
            if path:
                new_apps.append({"name":name,"path":path})
                if row["radio"].isChecked():
                    default_index=len(new_apps)-1
        if default_index==-1 and new_apps:
            default_index=0
        self.cronus_apps=new_apps
        self.cronus_default_index=default_index
        msgs=msg1+msg2
        if msgs:
            QMessageBox.information(self,"Parameters Adjusted","\n".join(msgs),QMessageBox.StandardButton.Ok)
        self.accept()
    def _add_app_entry(self):
        path,_=QFileDialog.getOpenFileName(self,"Select Cronus App (.exe or .lnk)",
                                           "", "Executable (*.exe);;Shortcut (*.lnk);;All Files (*)")
        if not path: return
        name=os.path.basename(path)
        self.cronus_apps.append({"name":name,"path":path})
        self._rebuild_apps_ui()
    def _rebuild_apps_ui(self):
        # Clear old
        while self.apps_container.count():
            item=self.apps_container.takeAt(0)
            w=item.widget()
            if w: w.deleteLater()
        self.app_rows=[]
        for i,app in enumerate(self.cronus_apps):
            row_frame=QFrame()
            row_layout=QHBoxLayout(row_frame); row_layout.setContentsMargins(4,4,4,4); row_layout.setSpacing(10)
            radio=QRadioButton()
            radio.setChecked(i==self.cronus_default_index)
            name_edit=QLineEdit(app["name"]); name_edit.setFixedWidth(160)
            path_edit=QLineEdit(app["path"]); path_edit.setReadOnly(True); path_edit.setMinimumWidth(340)
            browse_btn=QPushButton("Browse")
            browse_btn.setStyleSheet("""
                QPushButton {background:#3b82f6;color:white;border:none;padding:4px 10px;
                             border-radius:6px;font-size:11px;font-weight:600;}
                QPushButton:hover {background:#2563eb;}
            """)
            def change_path(idx=i):
                new_path,_=QFileDialog.getOpenFileName(self,"Change App Path",
                                                       os.path.dirname(path_edit.text()) if path_edit.text() else "",
                                                       "Executable (*.exe);;Shortcut (*.lnk);;All Files (*)")
                if new_path:
                    path_edit.setText(new_path)
            browse_btn.clicked.connect(change_path)
            remove_btn=QPushButton("Remove")
            remove_btn.setStyleSheet("""
                QPushButton {background:#ef4444;color:white;border:none;padding:4px 10px;
                             border-radius:6px;font-size:11px;font-weight:600;}
                QPushButton:hover {background:#dc2626;}
            """)
            def remove(idx=i):
                self.cronus_apps.pop(idx)
                if self.cronus_default_index==idx:
                    self.cronus_default_index=-1
                elif self.cronus_default_index>idx:
                    self.cronus_default_index-=1
                self._rebuild_apps_ui()
            remove_btn.clicked.connect(remove)
            row_layout.addWidget(radio)
            row_layout.addWidget(QLabel("Name:")); row_layout.addWidget(name_edit)
            row_layout.addWidget(QLabel("Path:")); row_layout.addWidget(path_edit)
            row_layout.addWidget(browse_btn); row_layout.addWidget(remove_btn); row_layout.addStretch()
            self.apps_container.addWidget(row_frame)
            self.app_rows.append({
                "radio":radio,"name_edit":name_edit,"path_edit":path_edit
            })
    def get_show_map(self): return self.map_checkbox.isChecked()
    def get_show_zones(self): return self.zones_checkbox.isChecked()
    def get_log_dir(self): return self.log_dir_input.text().strip()
    def get_ch_params(self,channel): return self.ch1_params if channel==1 else self.ch2_params
    def get_device_ranges(self): return self.dev_ranges
    def get_zones_cfg(self): return self.zones_cfg
    def get_cronus_apps(self): return self.cronus_apps
    def get_cronus_default_index(self): return self.cronus_default_index

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Cronus Training App")
        self.setMinimumSize(1300,800)
        self.config=load_config()
        self.device_ranges={1:fetch_device_range(1),2:fetch_device_range(2)}
        self._initialize_channel_params(self.config['ch1'],self.device_ranges[1])
        self._initialize_channel_params(self.config['ch2'],self.device_ranges[2])
        self.show_map=True
        self.show_zones=self.config.get("show_zones",False)
        self.zones_cfg=self.config.get("zones",[])
        self.cronus_apps=self.config.get("cronus_apps",[])
        self.cronus_default_index=self.config.get("cronus_app_default",-1)
        self._map_bars=[]
        self._apply_theme()
        self.wavelength_map_data={}
        self._setup_ui()
        if (self.device_ranges[1][0] is None or self.device_ranges[2][0] is None):
            QTimer.singleShot(300,self.open_settings)
        self.status_worker=StatusWorker()
        self.status_worker.status_update.connect(self.on_status_update)
        self.status_worker.start()
        self.map_dragging=False
        self._connect_map_events()
    def _initialize_channel_params(self,ch_params:dict,device_range:tuple):
        dmin,dmax=device_range
        if dmin is None or dmax is None: return
        if ch_params.get("test_min") is None or ch_params.get("test_max") is None:
            ch_params["test_min"]=dmin; ch_params["test_max"]=dmax
        ch_params["test_min"]=max(ch_params["test_min"],dmin)
        ch_params["test_max"]=min(ch_params["test_max"],dmax)
        if ch_params["test_min"]>=ch_params["test_max"]:
            ch_params["test_min"],ch_params["test_max"]=dmin,dmax
    def closeEvent(self,event):
        if hasattr(self,'status_worker'):
            self.status_worker.stop(); self.status_worker.wait()
        event.accept()
    def _apply_theme(self):
        self.setStyleSheet("""
            QWidget {background-color:#f2f6fa;color:#1e293b;font-family:'Segoe UI',Tahoma,sans-serif;}
            QScrollArea {border:none;background:transparent;}
        """)
    def _setup_ui(self):
        layout=QVBoxLayout(self); layout.setContentsMargins(18,18,18,18); layout.setSpacing(16)
        header=QFrame(); header.setStyleSheet("""
            QFrame {background:#ffffff;border:1px solid #e2e8f0;border-radius:22px;padding:6px;}
        """)
        h_l=QHBoxLayout(header); h_l.setContentsMargins(20,10,20,10)
        logo=QLabel()
        if os.path.exists("gaia_logo.png"):
            pm=QPixmap("gaia_logo.png").scaled(44,44,Qt.AspectRatioMode.KeepAspectRatio,
                                               Qt.TransformationMode.SmoothTransformation)
            logo.setPixmap(pm)
        h_l.addWidget(logo)
        title=QLabel("Cronus Training App"); title.setFont(QFont("Segoe UI",22,QFont.Weight.Bold))
        h_l.addWidget(title)
        # Cronus app selection combo
        self.cronus_app_combo=QComboBox()
        self._populate_cronus_combo()
        self.cronus_app_combo.setMinimumWidth(220)
        h_l.addWidget(self.cronus_app_combo)
        launch_btn=QPushButton("Launch Cronus App")
        launch_btn.setToolTip("Launch selected external Cronus application")
        launch_btn.clicked.connect(self.on_launch_cronus_app)
        launch_btn.setStyleSheet(self._header_button_style())
        h_l.addWidget(launch_btn)
        h_l.addStretch()
        status_frame=QFrame(); status_frame.setStyleSheet("""
            QFrame {background:#f8fafc;border:1px solid #e2e8f0;border-radius:14px;padding:4px 14px;}
        """)
        sf=QHBoxLayout(status_frame); sf.setContentsMargins(8,4,8,4)
        sf.addWidget(QLabel("Status:"))
        self.connection_indicator=QLabel("●"); self.connection_indicator.setStyleSheet("color:#94a3b8;font-size:16px;")
        self.connection_status=QLabel("Disconnected"); self.connection_status.setStyleSheet("font-weight:600;")
        sf.addWidget(self.connection_indicator); sf.addWidget(self.connection_status)
        sf.addSpacing(10); self.mode_label=QLabel("Mode: Unknown"); sf.addWidget(self.mode_label)
        h_l.addWidget(status_frame)
        reconnect_btn=QPushButton("Reconnect")
        reconnect_btn.setToolTip("Manually refresh connection & ranges")
        reconnect_btn.setStyleSheet(self._header_button_style())
        reconnect_btn.clicked.connect(self.on_reconnect)
        h_l.addWidget(reconnect_btn)
        settings_btn=QPushButton("Settings"); settings_btn.setStyleSheet(self._header_button_style())
        settings_btn.clicked.connect(self.open_settings); h_l.addWidget(settings_btn)
        self.shutdown_btn=QPushButton("Shutdown"); self.shutdown_btn.setStyleSheet(self._header_button_style(red=True))
        self.shutdown_btn.clicked.connect(self.on_shutdown_cronus); h_l.addWidget(self.shutdown_btn)
        layout.addWidget(header)
        self.map_frame=QFrame(); self.map_frame.setStyleSheet("""
            QFrame {background:#ffffff;border:1px solid #e2e8f0;border-radius:22px;}
        """)
        map_layout=QVBoxLayout(self.map_frame); map_layout.setContentsMargins(18,14,18,14)
        map_head=QHBoxLayout()
        self.map_title=QLabel("Wavelength Testing Map"); self.map_title.setFont(QFont("Segoe UI",16,QFont.Weight.Bold))
        map_head.addWidget(self.map_title)
        self.map_success_label=QLabel("")
        self.map_success_label.setStyleSheet("font-size:13px;font-weight:600;color:#334155;"
                                             "background:#e2e8f0;padding:4px 10px;border-radius:14px;")
        map_head.addSpacing(20); map_head.addWidget(self.map_success_label); map_head.addStretch()
        map_reset=QPushButton("Reset Map"); map_reset.setStyleSheet(self._header_button_style())
        map_reset.clicked.connect(self.on_reset_wavelength_map); map_head.addWidget(map_reset)
        map_layout.addLayout(map_head)
        self.map_fig,self.map_ax=plt.subplots(figsize=(10,2.6))
        self.map_canvas=FigureCanvas(self.map_fig); self.update_wavelength_map()
        map_layout.addWidget(self.map_canvas)
        layout.addWidget(self.map_frame); self.map_frame.setVisible(self.show_map)
        scroll=QScrollArea(); scroll.setWidgetResizable(True)
        inner=QWidget(); ch_l=QHBoxLayout(inner); ch_l.addStretch()
        self.ch1_panel=ChannelPanel(1,self.get_channel_params,self.get_device_range)
        self.ch2_panel=ChannelPanel(2,self.get_channel_params,self.get_device_range)
        self.ch1_panel.wavelength_tested=self.on_wavelength_tested
        self.ch2_panel.wavelength_tested=self.on_wavelength_tested
        ch_l.addWidget(self.ch1_panel); ch_l.addSpacing(32); ch_l.addWidget(self.ch2_panel); ch_l.addStretch()
        scroll.setWidget(inner); layout.addWidget(scroll)
    def _header_button_style(self,red=False):
        if red: base="#e74c3c"; hover="#c0392b"
        else: base="#34495e"; hover="#4d6070"
        return f"""
            QPushButton {{background:{base};color:white;border:none;padding:10px 20px;border-radius:14px;
                          font-weight:600;font-size:13px;}}
            QPushButton:hover {{background:{hover};}}
            QPushButton:disabled {{background:#b4c0cb;color:#e2e8f0;}}
        """
    def _populate_cronus_combo(self):
        self.cronus_app_combo.clear()
        for app in self.cronus_apps:
            self.cronus_app_combo.addItem(app["name"])
        if 0<=self.cronus_default_index<len(self.cronus_apps):
            self.cronus_app_combo.setCurrentIndex(self.cronus_default_index)
    def open_settings(self):
        dlg=SettingsDialog(self,self.show_map,self.show_zones,self.config['log_dir'],
                           self.config['ch1'],self.config['ch2'],self.device_ranges,
                           self.zones_cfg,self.cronus_apps,self.cronus_default_index)
        if dlg.exec()==QDialog.DialogCode.Accepted:
            self.show_map=dlg.get_show_map(); self.show_zones=dlg.get_show_zones()
            self.config['show_zones']=self.show_zones
            new_dir=dlg.get_log_dir()
            if new_dir and new_dir!=self.config['log_dir']:
                set_log_dir(new_dir); self.config['log_dir']=new_dir
            self.config['ch1']=dlg.get_ch_params(1); self.config['ch2']=dlg.get_ch_params(2)
            self.zones_cfg=dlg.get_zones_cfg(); self.config['zones']=self.zones_cfg
            self.cronus_apps=dlg.get_cronus_apps(); self.config['cronus_apps']=self.cronus_apps
            self.cronus_default_index=dlg.get_cronus_default_index(); self.config['cronus_app_default']=self.cronus_default_index
            self._initialize_channel_params(self.config['ch1'],self.device_ranges[1])
            self._initialize_channel_params(self.config['ch2'],self.device_ranges[2])
            save_config(self.config)
            self.map_frame.setVisible(self.show_map)
            self.ch1_panel.refresh_param_summary(); self.ch2_panel.refresh_param_summary()
            self.update_wavelength_map()
            self._populate_cronus_combo()
    def get_channel_params(self,ch): return self.config['ch1'] if ch==1 else self.config['ch2']
    def get_device_range(self,ch): return self.device_ranges.get(ch,(None,None))
    def on_status_update(self,connected,mode):
        if connected:
            self.connection_indicator.setStyleSheet("color:#22c55e;font-size:16px;")
            self.connection_status.setText("Connected"); self.shutdown_btn.setEnabled(True)
        else:
            self.connection_indicator.setStyleSheet("color:#ef4444;font-size:16px;")
            self.connection_status.setText("Disconnected"); self.shutdown_btn.setEnabled(False)
        self.mode_label.setText(f"Mode: {mode}")
    def on_shutdown_cronus(self):
        result=safe_put_json(f"{API_BASE}/Off",{})
        if result and result.get("OK"):
            self.connection_status.setText("Shutting down...")
            self.connection_indicator.setStyleSheet("color:#f39c12;font-size:16px;")
    def on_reconnect(self):
        refreshed=False
        for ch in (1,2):
            rng=fetch_device_range(ch)
            if rng[0] is not None and rng[1] is not None:
                self.device_ranges[ch]=rng; refreshed=True
        if refreshed:
            self._initialize_channel_params(self.config['ch1'],self.device_ranges[1])
            self._initialize_channel_params(self.config['ch2'],self.device_ranges[2])
            self.ch1_panel.refresh_param_summary(); self.ch2_panel.refresh_param_summary()
            save_config(self.config)
        status_response=safe_get_json(f"{API_BASE}/Status")
        connected=status_response is not None and status_response.get("OK",False)
        self.on_status_update(connected,status_response.get("Mode","Unknown") if connected else "Unknown")
        if refreshed and connected:
            QMessageBox.information(self,"Reconnect","Device ranges updated and connection active.",
                                    QMessageBox.StandardButton.Ok)
        elif refreshed:
            QMessageBox.information(self,"Reconnect","Ranges updated. Backend still disconnected.",
                                    QMessageBox.StandardButton.Ok)
        else:
            QMessageBox.information(self,"Reconnect","No ranges retrieved (backend may still be starting).",
                                    QMessageBox.StandardButton.Ok)
    def on_launch_cronus_app(self):
        if not self.cronus_apps:
            reply=QMessageBox.question(self,"Launch Cronus App",
                                       "No app entries configured. Open Settings now?",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                       QMessageBox.StandardButton.Yes)
            if reply==QMessageBox.StandardButton.Yes:
                self.open_settings()
            return
        idx=self.cronus_app_combo.currentIndex()
        if not (0<=idx<len(self.cronus_apps)):
            QMessageBox.warning(self,"Launch Cronus App","Invalid selection.")
            return
        path=self.cronus_apps[idx]["path"]
        if not os.path.exists(path):
            QMessageBox.warning(self,"Launch Cronus App",f"Path does not exist:\n{path}")
            return
        try:
            # Prefer os.startfile for .lnk or direct exe launching
            if path.lower().endswith(".lnk"):
                os.startfile(path)
            else:
                subprocess.Popen(f'"{path}"', shell=True)
            QMessageBox.information(self,"Launch Cronus App",f"Launched:\n{path}",
                                    QMessageBox.StandardButton.Ok)
        except Exception as e:
            QMessageBox.warning(self,"Launch Cronus App",f"Failed to launch:\n{e}")
    def on_wavelength_tested(self,wl,success):
        self.wavelength_map_data[wl]='success' if success else 'failed'
        self.update_wavelength_map()
    def on_reset_wavelength_map(self):
        self.wavelength_map_data.clear(); self.update_wavelength_map()
    def _connect_map_events(self):
        self.map_fig.canvas.mpl_connect("button_press_event",self._on_map_press)
        self.map_fig.canvas.mpl_connect("button_release_event",self._on_map_release)
        self.map_fig.canvas.mpl_connect("motion_notify_event",self._on_map_motion)
    def update_wavelength_map(self):
        if not hasattr(self,'_map_bars'): self._map_bars=[]
        ax=self.map_ax; ax.clear()
        bg_color='white'; text_color='#1e293b'; grid_color='#cbd5e1'
        data=self.wavelength_map_data; self._map_bars.clear()
        enabled_zones=[]
        for i,z in enumerate(self.zones_cfg):
            if i>=len(BASE_ZONE_DEFS): break
            if z.get("enabled",True):
                enabled_zones.append({
                    "label": z.get("name",f"Zone{i+1}"),
                    "min": float(z.get("min",BASE_ZONE_DEFS[i]["fixed_min"])),
                    "max": float(z.get("max",BASE_ZONE_DEFS[i]["fixed_max"])),
                    "color": BASE_ZONE_DEFS[i]["color"]
                })
        if data:
            tested=list(data.keys()); data_min=min(tested); data_max=max(tested)
        else:
            data_min=data_max=None
        if enabled_zones:
            zone_min=min(z['min'] for z in enabled_zones)
            zone_max=max(z['max'] for z in enabled_zones)
        else:
            zone_min,zone_max=670.0,900.0
        if data_min is None:
            x_min_display=zone_min if self.show_zones else 670.0
            x_max_display=zone_max if self.show_zones else 900.0
        else:
            x_min_display=min(data_min,zone_min) if self.show_zones else data_min
            x_max_display=max(data_max,zone_max) if self.show_zones else data_max
        if self.show_zones:
            for z in enabled_zones:
                ax.axvspan(z['min'],z['max'],ymin=0,ymax=1,facecolor=z['color'],alpha=0.22,zorder=0)
                mid=(z['min']+z['max'])/2
                ax.text(mid,1.02,z['label'],ha='center',va='bottom',fontsize=9,color='#334155',
                        fontweight='600',clip_on=False,alpha=0.9)
        if not data:
            ax.text(0.5,0.5,"No wavelengths tested yet",ha='center',va='center',fontsize=12,
                    color='#64748b',transform=ax.transAxes)
            self.map_success_label.setText("")
            self.map_success_label.setStyleSheet("font-size:13px;font-weight:600;color:#334155;"
                                                 "background:#e2e8f0;padding:4px 10px;border-radius:14px;")
        else:
            success_wls=[wl for wl,st in data.items() if st=='success']
            fail_wls=[wl for wl,st in data.items() if st=='failed']
            total=len(data); succ=len(success_wls); rate=(succ/total*100) if total else 0
            self.map_success_label.setText(f"{rate:.1f}% Success")
            if rate>95:
                pill_bg="#dcfce7"; pill_fg="#166534"; border="#16a34a"
            elif rate>=80:
                pill_bg="#fef9c3"; pill_fg="#854d0e"; border="#f59e0b"
            else:
                pill_bg="#fee2e2"; pill_fg="#7f1d1d"; border="#ef4444"
            self.map_success_label.setStyleSheet(
                f"font-size:13px;font-weight:600;color:{pill_fg};background:{pill_bg};"
                f"padding:4px 12px;border-radius:16px;border:1px solid {border};"
            )
            span=x_max_display - x_min_display if x_max_display>x_min_display else 100
            bar_w=max(0.5,min(span*0.015,5))
            if fail_wls:
                bf=ax.bar(fail_wls,[0.4]*len(fail_wls),width=bar_w,bottom=0,
                          color='#ef4444',edgecolor='#dc2626',linewidth=1.2,label='Failed',picker=True,zorder=5)
                for b,wl in zip(bf,fail_wls): self._map_bars.append((b,wl,'Failed'))
            if success_wls:
                bs=ax.bar(success_wls,[0.4]*len(success_wls),width=bar_w,bottom=0.5,
                          color='#22c55e',edgecolor='#16a34a',linewidth=1.2,label='Success',picker=True,zorder=5)
                for b,wl in zip(bs,success_wls): self._map_bars.append((b,wl,'Success'))
        final_span=x_max_display - x_min_display if x_max_display>x_min_display else 100
        pad=max(final_span*0.05,10)
        ax.set_xlim(x_min_display - pad, x_max_display + pad); ax.set_ylim(0,1)
        ax.set_xlabel("Wavelength (nm)",fontsize=11,fontweight='600',color=text_color)
        ax.set_yticks([])
        for side in ['top','right','left']: ax.spines[side].set_visible(False)
        ax.spines['bottom'].set_color(grid_color); ax.spines['bottom'].set_linewidth(1.4)
        ax.tick_params(axis='x',colors=text_color,labelsize=9)
        ax.grid(True,axis='x',linestyle='--',alpha=0.18,color=grid_color); ax.set_facecolor(bg_color)
        if data:
            legend=self.map_fig.legend(loc='lower center',ncol=2,framealpha=0.98,
                                       facecolor=bg_color,edgecolor=grid_color,
                                       bbox_to_anchor=(0.5,-0.18),fontsize=10,
                                       borderpad=1,labelspacing=0.8,handlelength=2)
            for txt in legend.get_texts():
                txt.set_color(text_color); txt.set_fontweight('600')
        else:
            if self.map_fig.legends:
                for L in self.map_fig.legends: L.remove()
        self.map_fig.subplots_adjust(bottom=0.30); self.map_canvas.draw()
    def _on_map_press(self,event):
        if event.inaxes==self.map_ax and event.button==1: self.map_dragging=True
    def _on_map_release(self,event):
        self.map_dragging=False; QToolTip.hideText()
    def _on_map_motion(self,event):
        if not self.map_dragging: return
        if event.inaxes!=self.map_ax or event.xdata is None or event.ydata is None:
            QToolTip.hideText(); return
        for bar,wl,status in self._map_bars:
            bx=bar.get_x(); bw=bar.get_width(); by=bar.get_y(); bh=bar.get_height()
            if bx<=event.xdata<=bx+bw and by<=event.ydata<=by+bh:
                text=f"{wl:.1f} nm - {'✓' if status=='Success' else '✗'} {status}"
                global_pos=self.map_canvas.mapToGlobal(
                    QPoint(int(event.guiEvent.position().x()),int(event.guiEvent.position().y())))
                QToolTip.showText(global_pos+QPoint(12,12),text,self.map_canvas,msecShowTime=1200)
                return
        QToolTip.hideText()

def main():
    app=QApplication(sys.argv)
    win=MainWindow()
    win.showMaximized()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()