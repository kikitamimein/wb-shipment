import sys
import os
import csv
import math
from collections import defaultdict
import pandas as pd

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QStackedWidget, QTabWidget,
    QTableWidget, QTableWidgetItem, QHeaderView, QCheckBox, QSpinBox,
    QMessageBox, QLineEdit, QFormLayout, QGridLayout,
    QDialog, QScrollArea, QDialogButtonBox, QComboBox, QProgressBar, QMenu, QWidgetAction,
    QListWidget, QListWidgetItem, QAbstractItemView, QRadioButton
)
from PyQt6.QtCore import Qt, pyqtSignal, QSettings, QThread
from PyQt6.QtGui import QFont, QIcon, QDropEvent, QDragEnterEvent, QColor, QBrush, QKeySequence, QAction, QMouseEvent

import wb_api
import moysklad_api
import calculator
from texts import TEXTS

from qt_material import apply_stylesheet, list_themes

APP_COMPANY = "MyKompany"
APP_NAME = "WBtoMoySklad"

DEFAULT_WB1_KEY = "eyJhbGciOiJFUzI1NiIsImtpZCI6IjIwMjYwMzAydjEiLCJ0eXAiOiJKV1QifQ.eyJhY2MiOjEsImVudCI6MSwiZXhwIjoxNzg5Njk5MTg4LCJpZCI6IjAxOWQwNjg5LWQyYjEtNzMwZi05ODRlLWVmMDkxZmJlNDExZiIsImlpZCI6NDk1MTc4MDUsIm9pZCI6MjUxNDk3LCJzIjoxMDYwLCJzaWQiOiIwYzVhZGM3Ny0wNGFhLTQ3NjctYWY4NC01ODczMTY3YmU5Y2QiLCJ0IjpmYWxzZSwidWlkIjo0OTUxNzgwNX0.AclK9wZ16h_kj8d-dpC0vnTGxTkH1FutZnRnZBhVe6SdD48avi3ijwaNQvxUuiherzzXaBEz-MmDcRWTxG7H0Q"
DEFAULT_WB2_KEY = "eyJhbGciOiJFUzI1NiIsImtpZCI6IjIwMjYwMzAydjEiLCJ0eXAiOiJKV1QifQ.eyJhY2MiOjEsImVudCI6MSwiZXhwIjoxNzkxMzI3NDAxLCJpZCI6IjAxOWQ2Nzk2LTVmZTUtN2U3ZC1iNmU2LWYwZDhhM2UzMDA5YyIsImlpZCI6MjY0NjU5OTAsIm9pZCI6NDM3Nzk5NSwicyI6MTA2MCwic2lkIjoiYjNiODg5MjktNjViMy00YjFhLTk3ZjYtZTg4MGQ4M2ViOWQ3IiwidCI6ZmFsc2UsInVpZCI6MjY0NjU5OTB9.NWuKXQicIcg5kmhLL_qacyeBPmJw08Uv2-YVlx_kCOPkQpvWCzDQ1dSJfIQh-74l9ezdzXH-8SRl_j5UiFq2gw"
DEFAULT_MS_TOKEN = "91ba131f32266b3baaf51eef0215f44d02350631"

class CustomTableWidget(QTableWidget):
    def mousePressEvent(self, event):
        item = self.itemAt(event.pos())
        if not item:
            self.clearSelection()
        super().mousePressEvent(event)
        
    def keyPressEvent(self, event):
        if event.matches(QKeySequence.StandardKey.Copy):
            selected = self.selectedRanges()
            if selected:
                lines = []
                r = selected[0]
                for i in range(r.topRow(), r.bottomRow() + 1):
                    row_data = []
                    for j in range(r.leftColumn(), r.rightColumn() + 1):
                        item = self.item(i, j)
                        row_data.append(item.text().strip() if item else "")
                    lines.append("\t".join(row_data))
                text = "\n".join(lines)
                QApplication.clipboard().setText(text)
        else:
            super().keyPressEvent(event)

class NumericTableWidgetItem(QTableWidgetItem):
    def __lt__(self, other):
        try:
            v1 = float(self.data(Qt.ItemDataRole.UserRole))
            v2 = float(other.data(Qt.ItemDataRole.UserRole))
            return v1 < v2
        except (ValueError, TypeError):
            return super().__lt__(other)

class FinalShipmentWidget(QWidget):
    valueChanged = pyqtSignal(int)
    skipChanged = pyqtSignal(bool)
    
    def __init__(self, value, is_skipped=False):
        super().__init__()
        layout = QHBoxLayout(self)
        layout.setContentsMargins(5, 2, 5, 2)
        layout.setSpacing(5)
        
        self.btn = QPushButton("✖")
        self.btn.setCheckable(True)
        self.btn.setFixedSize(18, 18)
        self.btn.setChecked(is_skipped)
        self.update_btn(is_skipped)
        self.btn.toggled.connect(self.on_toggled)
        self.btn.setToolTip("Исключить из отгрузки (без уменьшения остатков МС)")
        self.btn.setCursor(Qt.CursorShape.PointingHandCursor)
        
        self.spin = QSpinBox()
        self.spin.setRange(0, 100000)
        self.spin.setValue(int(value))
        self.spin.setStyleSheet("QSpinBox { background: transparent; border: none; font-weight: bold; color: white; } QSpinBox::up-button, QSpinBox::down-button { width: 0px; }")
        self.spin.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.spin.valueChanged.connect(self.valueChanged.emit)
        
        layout.addStretch()
        layout.addWidget(self.spin)
        layout.addStretch()
        layout.addWidget(self.btn)
        layout.setAlignment(self.btn, Qt.AlignmentFlag.AlignVCenter)
        
    def on_toggled(self, checked):
        self.update_btn(checked)
        self.skipChanged.emit(checked)
        
    def update_btn(self, checked):
        if checked:
            self.btn.setStyleSheet("QPushButton { background-color: #ff4d4d; color: white; border-radius: 4px; font-weight: bold; }")
        else:
            self.btn.setStyleSheet("QPushButton { background-color: transparent; color: #888; border-radius: 4px; border: 1px solid #555; }")

class AppState:
    def __init__(self):
        self.settings = QSettings(APP_COMPANY, APP_NAME)
        
        self.excel_path = ""
        self.stocks_data = []    # Raw stocks from excel
        self.warehouses = []     # List of warehouse names
        self.active_warehouses = []
        self.global_turnovers = {} 
        self.custom_turnovers = {} # {warehouseName: {supplierArticle: turnover_days}}
        self.manual_overrides = {} # {warehouseName: {supplierArticle: qty}}
        self.calculated_data = {}  # {warehouseName: list of dicts}
        self.ms_stocks_details = {}
        self.ms_stocks = {}
        self.ms_reserved_local = {}
        self.custom_rows = {}      # {warehouseName: list of dicts {"supplierArticle": art, "quantity": qty}}
        self.excluded_articles = {}# {warehouseName: set(arts)}
        self.skipped_articles = {} # {warehouseName: set(arts)}
        self.export_warehouses = set() # {warehouseName}
        

        
        self.search_query = ""
        
        # Load column visibility preferences
        self.visible_columns = {
            "subject": self.get_setting("col_subject", True),
            "name": self.get_setting("col_name", True),
            "stock": self.get_setting("col_stock", True),
            "sales": self.get_setting("col_sales", True),
            "avg": self.get_setting("col_avg", True),
            "wb_turnover": self.get_setting("col_wb_turnover", True),
            "turnover": self.get_setting("col_turnover", True),
            "target": self.get_setting("col_target", True),
            "need": self.get_setting("col_need", True),
            "final": self.get_setting("col_final", True),
            "ms_stock": self.get_setting("col_ms_stock", True)
        }
        
        self.rounding_to_5 = self.get_setting("rounding_to_5", False)
        self.hide_zero = self.get_setting("hide_zero", False)
        self.hide_zero_ms = self.get_setting("hide_zero_ms", False)
        self.limit_to_stock = self.get_setting("limit_to_stock", False)
        self.turnover_filter_enabled = self.get_setting("turnover_filter_enabled", False)
        self.turnover_filter_from = int(self.get_setting("turnover_filter_from", 0))
        self.turnover_filter_to = int(self.get_setting("turnover_filter_to", 9999))
        self.auto_hide_empty = self.get_setting("auto_hide_empty", False)
        self.use_missing_days = self.get_setting("use_missing_days", False)
        self.highlight_recent = self.get_setting("highlight_recent", False)
        self.recent_days = int(self.get_setting("recent_days", 14))
        self.recent_supplies = {}  # {article: received_date}, loaded on demand
        self.recent_supplies_detailed = []  # [dict: article, name, barcode, received_date, quantity, meta]
        self.recently_added = {}  # {warehouseName: set(articles)} — added via RecentSuppliesDialog
        self.theme = self.get_setting("theme", "dark_teal.xml")
        
        self.current_warehouse = ""

    def get_setting(self, key, default_val):
        val = self.settings.value(key)
        if val is None:
            return default_val
        # QSettings stores bools as "true"/"false" strings
        if isinstance(default_val, bool):
            return str(val).lower() == 'true' if isinstance(val, str) else bool(val)
        return val

    def set_setting(self, key, val):
        self.settings.setValue(key, val)

state = AppState()

def _is_within_days(date_str, days):
    """Check if date_str (ISO-like '2026-05-15 12:26:00') is within the last `days` days."""
    from datetime import datetime, timedelta
    try:
        d = datetime.strptime(date_str[:19], "%Y-%m-%d %H:%M:%S")
        return d >= datetime.now() - timedelta(days=days)
    except (ValueError, TypeError):
        return False

def calculate_all_warehouses():
    state.calculated_data.clear()
    state.ms_reserved_local.clear()
    
    all_needs = {}
    
    # 1. Calculate raw unconstrained needs for all warehouses
    for w in state.active_warehouses:
        w_stocks = [s for s in state.stocks_data if s["warehouseName"] == w]
        
        global_t = state.global_turnovers.get(w, 30)
        custom_t = state.custom_turnovers.get(w, {})
        
        needs = calculator.calculate_warehouse_needs(
            w_stocks, 
            default_turnover_days=global_t, 
            custom_turnovers=custom_t,
            round_to_5=state.rounding_to_5,
            use_missing_days=getattr(state, "use_missing_days", False)
        )
        
        filtered_needs = []
        excluded = state.excluded_articles.get(w, set())
        for item in needs:
            if str(item["supplierArticle"]) not in excluded:
                filtered_needs.append(item)
        needs = filtered_needs

        existing_arts = set(str(item["supplierArticle"]) for item in needs)
        for crow in state.custom_rows.get(w, []):
            art = str(crow["supplierArticle"])
            if art in excluded or art in existing_arts:
                continue
            qty = crow["quantity"]
            needs.append({
                "itemName": TEXTS["item_manual"],
                "supplierArticle": art,
                "target_stock": qty,
                "need": qty,
                "final_shipment": qty,
                "wb_turnover": 0
            })

        overrides = state.manual_overrides.get(w, {})
        skipped = state.skipped_articles.get(w, set())
        for item in needs:
            art = str(item["supplierArticle"])
            final_ship = int(item["final_shipment"])
            item["is_skipped"] = (art in skipped)
            
            base_ms = state.ms_stocks.get(art, 0)
            reserved = state.ms_reserved_local.get(art, 0)
            remain_ms = base_ms - reserved
            
            if state.limit_to_stock and remain_ms < final_ship:
                item["final_shipment"] = max(0, int(remain_ms))
                
            # Manual user overrides take absolute priority over constraints
            if art in overrides:
                item["final_shipment"] = overrides[art]
                
            if w in state.export_warehouses and not item.get("is_skipped", False):
                state.ms_reserved_local[art] = state.ms_reserved_local.get(art, 0) + int(item["final_shipment"])
                
        state.calculated_data[w] = needs


class DropZoneLabel(QLabel):
    file_dropped = pyqtSignal(str)
    
    def __init__(self, text):
        super().__init__(text)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setAcceptDrops(True)
        self.setStyleSheet("""
            QLabel {
                border: 2px dashed #888888;
                border-radius: 10px;
                background-color: transparent;
                padding: 40px;
                font-size: 16px;
                color: #aaaaaa;
            }
            QLabel:hover {
                border-color: #bbbbbb;
                background-color: rgba(255, 255, 255, 0.05);
            }
        """)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        if files:
            self.file_dropped.emit(files[0])


class DataLoaderThread(QThread):
    progress = pyqtSignal(str)
    finished_success = pyqtSignal()
    finished_error = pyqtSignal(str)

    def __init__(self, excel_path):
        super().__init__()
        self.excel_path = excel_path

    def run(self):
        try:
            # Setup dynamic API keys
            acc = getattr(state, "current_wb_account", "WB1")
            import wb_api
            import moysklad_api
            wb_api.WB_API_KEY_STATS = state.get_setting("WB1_API_KEY", DEFAULT_WB1_KEY) if acc == "WB1" else state.get_setting("WB2_API_KEY", DEFAULT_WB2_KEY)
            moysklad_api.MS_API_TOKEN = state.get_setting("MS_API_TOKEN", DEFAULT_MS_TOKEN)
            moysklad_api.MS_WB_ACCOUNT = acc
            
            self.progress.emit("1/4 Чтение таблицы Excel...")
            stocks = wb_api.get_stocks_from_excel(self.excel_path)
            state.stocks_data = stocks
            
            self.progress.emit("2/4 Получение данных из Моего Склада...")
            try:
                target_folder = "магазин на wildberries" if acc == "WB1" else "магазин на wildberries юля"
                state.ms_stocks_details = moysklad_api.get_all_stocks()
                
                state.ms_stocks = {}
                state.ms_folders = {}
                for k, v in state.ms_stocks_details.items():
                    folder = v.get("folder", "Без группы").strip().lower()
                    if folder == target_folder.lower():
                        state.ms_stocks[str(k)] = v["stock"]
                        state.ms_folders[str(k)] = v.get("folder")
                        
            except Exception as e:
                print(f"Error fetching MS stocks: {e}")

            # Pre-load recent supplies (60 days) once, client-side filtering later
            try:
                state.recent_supplies = moysklad_api.get_recent_supplies(days_back=60)
                state.recent_supplies_detailed = moysklad_api.get_recent_supplies_detailed(days_back=60)
            except Exception as e:
                print(f"Error fetching recent supplies: {e}")
                state.recent_supplies = {}
                state.recent_supplies_detailed = []

            state.ms_reserved_local.clear()
            
            self.progress.emit("3/4 Инициализация складов и анализ данных...")
            wh_totals = {}
            for s in stocks:
                wh_totals[s["warehouseName"]] = wh_totals.get(s["warehouseName"], 0) + s["quantity"]
                
            new_warehouses = sorted(list(wh_totals.keys()), key=lambda w: wh_totals[w], reverse=True)
            saved_order = state.get_setting("warehouses_order", [])
            
            ordered = []
            if isinstance(saved_order, list):
                for w in list(saved_order):
                    if w in new_warehouses:
                        ordered.append(w)
            for w in new_warehouses:
                if w not in ordered:
                    ordered.append(w)
                    
            state.warehouses = ordered
            
            saved_active = state.get_setting("active_warehouses", state.warehouses)
            if not isinstance(saved_active, (list, tuple)):
                saved_active = [saved_active] if saved_active else []
                
            state.active_warehouses = [w for w in state.warehouses if w in saved_active]
            
            if state.auto_hide_empty:
                for w in state.warehouses:
                    w_stocks = [s for s in state.stocks_data if s["warehouseName"] == w]
                    if sum(s.get("quantity", 0) for s in w_stocks) == 0:
                        if w in state.active_warehouses:
                            state.active_warehouses.remove(w)
                            
            if not state.active_warehouses:
                state.active_warehouses = state.warehouses.copy()
            
            for w in state.warehouses:
                try:
                    saved_t = int(state.get_setting(f"global_turnovers_{w}", 30))
                except:
                    saved_t = 30
                state.global_turnovers[w] = saved_t
                state.custom_turnovers[w] = {}
                state.manual_overrides[w] = {}
            
            calculate_all_warehouses()
            
            self.finished_success.emit()
            
        except Exception as e:
            self.finished_error.emit(str(e))


class ApiSettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Настройки API")
        self.resize(500, 200)
        
        layout = QFormLayout()
        
        self.ms_input = QLineEdit()
        self.ms_input.setText(state.get_setting("MS_API_TOKEN", DEFAULT_MS_TOKEN))
        
        self.wb1_input = QLineEdit()
        self.wb1_input.setText(state.get_setting("WB1_API_KEY", DEFAULT_WB1_KEY))
        
        self.wb2_input = QLineEdit()
        self.wb2_input.setText(state.get_setting("WB2_API_KEY", DEFAULT_WB2_KEY))
        
        layout.addRow("MS", self.ms_input)
        layout.addRow("WB1", self.wb1_input)
        layout.addRow("WB2", self.wb2_input)
        
        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn_box.accepted.connect(self.save_and_accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)
        
        self.setLayout(layout)

    def save_and_accept(self):
        state.set_setting("MS_API_TOKEN", self.ms_input.text().strip())
        state.set_setting("WB1_API_KEY", self.wb1_input.text().strip())
        state.set_setting("WB2_API_KEY", self.wb2_input.text().strip())
        self.accept()

class StartScreen(QWidget):
    def __init__(self, parent_nav):
        super().__init__()
        self.parent_nav = parent_nav
        
        layout = QVBoxLayout()
        
        top_layout = QHBoxLayout()
        top_layout.addStretch()
        self.btn_api_settings = QPushButton("⚙ Настройки API")
        self.btn_api_settings.clicked.connect(self.open_api_settings)
        top_layout.addWidget(self.btn_api_settings)
        layout.addLayout(top_layout)
        
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        title = QLabel(TEXTS["step1_title"])
        title.setFont(QFont("Arial", 20, QFont.Weight.Bold))
        layout.addWidget(title, alignment=Qt.AlignmentFlag.AlignCenter)
        
        desc = QLabel(TEXTS["step1_desc"])
        layout.addWidget(desc, alignment=Qt.AlignmentFlag.AlignCenter)
        
        self.drop_zone = DropZoneLabel(TEXTS["drop_zone"])
        self.drop_zone.file_dropped.connect(self.on_file_dropped)
        layout.addWidget(self.drop_zone)
        
        self.path_input = QLineEdit()
        self.path_input.setPlaceholderText(TEXTS["path_input_ph"])
        layout.addWidget(self.path_input)
        
        radio_layout = QHBoxLayout()
        self.radio_wb1 = QRadioButton("WB1")
        self.radio_wb2 = QRadioButton("WB2")
        self.radio_wb1.setChecked(True)
        radio_layout.addStretch()
        radio_layout.addWidget(self.radio_wb1)
        radio_layout.addWidget(self.radio_wb2)
        radio_layout.addStretch()
        layout.addLayout(radio_layout)
        
        btn_layout = QHBoxLayout()
        self.btn_pick = QPushButton(TEXTS["btn_pick"])
        self.btn_pick.clicked.connect(self.pick_file)
        
        self.btn_load = QPushButton(TEXTS["btn_load"])
        self.btn_load.clicked.connect(self.load_data)
        self.btn_load.setProperty('class', 'success')
        
        btn_layout.addWidget(self.btn_pick)
        btn_layout.addWidget(self.btn_load)
        layout.addLayout(btn_layout)
        
        # Loading indicators
        self.loading_label = QLabel("")
        self.loading_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.loading_label.hide()
        layout.addWidget(self.loading_label)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0) # Indeterminate mode -> bouncing dots / line
        self.progress_bar.hide()
        layout.addWidget(self.progress_bar)
        
        self.setLayout(layout)

    def open_api_settings(self):
        dlg = ApiSettingsDialog(self)
        dlg.exec()

    def on_file_dropped(self, filepath):
        self.path_input.setText(filepath)

    def pick_file(self):
        path, _ = QFileDialog.getOpenFileName(self, TEXTS["file_dialog_title"], "", "Excel Files (*.xlsx)")
        if path:
            self.path_input.setText(path)

    def load_data(self):
        path = self.path_input.text().strip()
        if not path or not os.path.exists(path):
            QMessageBox.warning(self, TEXTS["error_title"], TEXTS["file_not_found"])
            return
            
        state.excel_path = path
        state.current_wb_account = "WB1" if self.radio_wb1.isChecked() else "WB2"
        
        self.btn_load.setEnabled(False)
        self.btn_pick.setEnabled(False)
        self.path_input.setEnabled(False)
        self.drop_zone.setEnabled(False)
        
        self.loading_label.setText("Запуск обработки...")
        self.loading_label.show()
        self.progress_bar.show()
        
        # Start background processing thread
        self.loader_thread = DataLoaderThread(state.excel_path)
        self.loader_thread.progress.connect(self.on_loading_progress)
        self.loader_thread.finished_success.connect(self.on_loading_success)
        self.loader_thread.finished_error.connect(self.on_loading_error)
        self.loader_thread.start()

    def on_loading_progress(self, text):
        self.loading_label.setText(text)

    def on_loading_success(self):
        self._reset_loading_ui()
        self.parent_nav.go_to_warehouses()
        
    def on_loading_error(self, err_text):
        self._reset_loading_ui()
        QMessageBox.critical(self, TEXTS["error_loading_title"], TEXTS["error_loading"].format(e=err_text))
        
    def _reset_loading_ui(self):
        self.loading_label.hide()
        self.progress_bar.hide()
        self.btn_load.setEnabled(True)
        self.btn_pick.setEnabled(True)
        self.path_input.setEnabled(True)
        self.drop_zone.setEnabled(True)

class WarehouseTabWidget(QWidget):
    def __init__(self, warehouse_name):
        super().__init__()
        self.wh = warehouse_name
        self.is_populating = False
        
        layout = QVBoxLayout()
        
        control_layout = QHBoxLayout()
        
        left_layout = QVBoxLayout()
        left_layout.setSpacing(10)
        
        # Row 1: Search + Turnover + Table Settings
        row1_layout = QHBoxLayout()
        
        row1_layout.addSpacing(5)
        
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText(TEXTS["search_ph"])
        self.search_input.setMinimumWidth(280)
        self.search_input.textChanged.connect(self.refresh_table)
        row1_layout.addWidget(self.search_input, 1)
        
        row1_layout.addSpacing(15)
        
        row1_layout.addWidget(QLabel(TEXTS["turnover_lbl"]))
        self.global_t_spin = QSpinBox()
        self.global_t_spin.setRange(1, 365)
        self.global_t_spin.setMaximumWidth(60)
        self.global_t_spin.setValue(state.global_turnovers[self.wh])
        self.global_t_spin.valueChanged.connect(self.on_global_t_changed)
        row1_layout.addWidget(self.global_t_spin)
        
        row1_layout.addSpacing(20)
        
        self.rounding_cb = QCheckBox(TEXTS["cb_rounding"])
        self.rounding_cb.setChecked(state.rounding_to_5)
        self.rounding_cb.stateChanged.connect(self.on_rounding_changed)
        row1_layout.addWidget(self.rounding_cb)

        self.hide_zero_cb = QCheckBox(TEXTS["cb_hide_zero"])
        self.hide_zero_cb.setChecked(state.hide_zero)
        self.hide_zero_cb.stateChanged.connect(self.on_hide_zero_changed)
        row1_layout.addWidget(self.hide_zero_cb)

        self.hide_zero_ms_cb = QCheckBox(TEXTS.get("cb_hide_zero_ms", "Убрать 0 в МС"))
        self.hide_zero_ms_cb.setChecked(state.hide_zero_ms)
        self.hide_zero_ms_cb.stateChanged.connect(self.on_hide_zero_ms_changed)
        row1_layout.addWidget(self.hide_zero_ms_cb)

        self.limit_to_stock_cb = QCheckBox(TEXTS["cb_limit_ms"])
        self.limit_to_stock_cb.setChecked(state.limit_to_stock)
        self.limit_to_stock_cb.stateChanged.connect(self.on_limit_to_stock_changed)
        row1_layout.addWidget(self.limit_to_stock_cb)

        row1_layout.addSpacing(10)

        self.turnover_filter_cb = QCheckBox(TEXTS["cb_turnover_filter"])
        self.turnover_filter_cb.setChecked(state.turnover_filter_enabled)
        self.turnover_filter_cb.stateChanged.connect(self.on_turnover_filter_changed)
        row1_layout.addWidget(self.turnover_filter_cb)

        self.turnover_from = QSpinBox()
        self.turnover_from.setRange(0, 9999)
        self.turnover_from.setValue(state.turnover_filter_from)
        self.turnover_from.setMinimumWidth(65)
        self.turnover_from.setEnabled(state.turnover_filter_enabled)
        self.turnover_from.valueChanged.connect(self.on_turnover_filter_changed)
        row1_layout.addWidget(self.turnover_from)

        row1_layout.addWidget(QLabel("—"))

        self.turnover_to = QSpinBox()
        self.turnover_to.setRange(0, 9999)
        self.turnover_to.setValue(state.turnover_filter_to)
        self.turnover_to.setMinimumWidth(65)
        self.turnover_to.setEnabled(state.turnover_filter_enabled)
        self.turnover_to.valueChanged.connect(self.on_turnover_filter_changed)
        row1_layout.addWidget(self.turnover_to)

        row1_layout.addSpacing(10)

        self.highlight_recent_cb = QCheckBox(TEXTS["cb_highlight_recent"])
        self.highlight_recent_cb.setChecked(state.highlight_recent)
        self.highlight_recent_cb.stateChanged.connect(self.on_highlight_recent_changed)
        row1_layout.addWidget(self.highlight_recent_cb)

        self.recent_days = QSpinBox()
        self.recent_days.setRange(1, 60)
        self.recent_days.setValue(state.recent_days)
        self.recent_days.setMinimumWidth(55)
        self.recent_days.valueChanged.connect(self.on_recent_days_changed)
        row1_layout.addWidget(self.recent_days)
        row1_layout.addWidget(QLabel(TEXTS["recent_days_suffix"]))
        
        row1_layout.addStretch()
        left_layout.addLayout(row1_layout)
        
        control_layout.addLayout(left_layout)
        
        layout.addLayout(control_layout)
        
        self.table = CustomTableWidget()
        self.table.setSortingEnabled(False) # Sorting manually via right click
        self.table.setAlternatingRowColors(True)
        
        header = self.table.horizontalHeader()
        header.setSectionsMovable(True)
        header.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        header.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        
        # Left click -> highlight column
        header.sectionClicked.connect(self.table.selectColumn)
        # Right click -> sort column
        header.customContextMenuRequested.connect(self.on_header_context_menu)
        
        self.table.itemChanged.connect(self.on_item_changed)
        layout.addWidget(self.table)
        
        row_tools_layout = QHBoxLayout()
        btn_add = QPushButton(TEXTS["btn_add_row"])
        btn_rem = QPushButton(TEXTS["btn_rem_row"])
        self.btn_undo = QPushButton(TEXTS["btn_undo"])
        self.btn_redo = QPushButton(TEXTS["btn_redo"])
        
        btn_add.clicked.connect(self.add_row)
        btn_rem.clicked.connect(self.rem_row)
        self.btn_undo.clicked.connect(self.undo)
        self.btn_redo.clicked.connect(self.redo)
        
        self.cb_add_to_export = QCheckBox(TEXTS.get("cb_add_to_export", "Добавить в итог"))
        self.cb_add_to_export.setChecked(self.wh in state.export_warehouses)
        self.cb_add_to_export.stateChanged.connect(self.on_add_to_export_changed)
        
        row_tools_layout.addWidget(btn_add)
        row_tools_layout.addWidget(btn_rem)
        row_tools_layout.addWidget(self.btn_undo)
        row_tools_layout.addWidget(self.btn_redo)
        row_tools_layout.addStretch()
        row_tools_layout.addWidget(self.cb_add_to_export)
        layout.addLayout(row_tools_layout)
        
        self.history = []
        self.history_idx = -1
        
        self.setLayout(layout)
        self.setup_table()
        self.save_state()

        # Restore column widths/order if they exist in QSettings for this specific table configuration
        saved_state = state.settings.value("tableHeaderState")
        if saved_state:
            self.table.horizontalHeader().restoreState(saved_state)

    def on_global_t_changed(self, val):
        state.global_turnovers[self.wh] = val
        state.set_setting(f"global_turnovers_{self.wh}", val)
        calculate_all_warehouses()
        self.refresh_table()

    def on_rounding_changed(self, cb_state):
        state.rounding_to_5 = cb_state == Qt.CheckState.Checked.value
        state.set_setting("rounding_to_5", state.rounding_to_5) # Persistent save
        calculate_all_warehouses()
        self.refresh_table()

    def on_hide_zero_changed(self, cb_state):
        state.hide_zero = cb_state == Qt.CheckState.Checked.value
        state.set_setting("hide_zero", state.hide_zero) # Persistent save
        self.refresh_table()

    def on_hide_zero_ms_changed(self, cb_state):
        state.hide_zero_ms = cb_state == Qt.CheckState.Checked.value
        state.set_setting("hide_zero_ms", state.hide_zero_ms)
        self.refresh_table()

    def on_limit_to_stock_changed(self, cb_state):
        state.limit_to_stock = cb_state == Qt.CheckState.Checked.value
        state.set_setting("limit_to_stock", state.limit_to_stock)
        calculate_all_warehouses()
        self.refresh_table()

    def on_turnover_filter_changed(self):
        enabled = self.turnover_filter_cb.isChecked()
        state.turnover_filter_enabled = enabled
        state.turnover_filter_from = self.turnover_from.value()
        state.turnover_filter_to = self.turnover_to.value()
        state.set_setting("turnover_filter_enabled", enabled)
        state.set_setting("turnover_filter_from", state.turnover_filter_from)
        state.set_setting("turnover_filter_to", state.turnover_filter_to)
        self.turnover_from.setEnabled(enabled)
        self.turnover_to.setEnabled(enabled)
        self.refresh_table()

    def on_highlight_recent_changed(self, cb_state):
        enabled = cb_state == Qt.CheckState.Checked.value
        state.highlight_recent = enabled
        state.set_setting("highlight_recent", enabled)

        if enabled and state.recent_supplies:
            parent = self.window()
            if parent:
                parent.statusBar().showMessage(
                    f"Загружено поступлений за 60 дн: {len(state.recent_supplies)} товаров", 4000)
        elif enabled:
            parent = self.window()
            if parent:
                parent.statusBar().showMessage(
                    "Нет данных о поступлениях — перезагрузите Excel-файл", 5000)

        self.refresh_table()

    def on_recent_days_changed(self, val):
        state.recent_days = val
        state.set_setting("recent_days", val)
        if state.highlight_recent:
            self.refresh_table()

    def update_undo_redo_buttons(self):
        self.btn_undo.setEnabled(self.history_idx > 0)
        self.btn_redo.setEnabled(self.history_idx < len(self.history) - 1)

    def save_state(self):
        import copy
        current = {
            "custom_rows": copy.deepcopy(state.custom_rows.get(self.wh, [])),
            "excluded": copy.deepcopy(state.excluded_articles.get(self.wh, set())),
            "skipped": copy.deepcopy(state.skipped_articles.get(self.wh, set())),
            "overrides": copy.deepcopy(state.manual_overrides.get(self.wh, {})),
            "turnovers": copy.deepcopy(state.custom_turnovers.get(self.wh, {}))
        }
        if self.history_idx < len(self.history) - 1:
            self.history = self.history[:self.history_idx + 1]
            
        if self.history and self.history[-1] == current:
            return
            
        self.history.append(current)
        self.history_idx += 1
        self.update_undo_redo_buttons()

    def undo(self):
        if self.history_idx > 0:
            self.history_idx -= 1
            self.restore_state(self.history[self.history_idx])

    def redo(self):
        if self.history_idx < len(self.history) - 1:
            self.history_idx += 1
            self.restore_state(self.history[self.history_idx])

    def restore_state(self, s):
        import copy
        state.custom_rows[self.wh] = copy.deepcopy(s["custom_rows"])
        state.excluded_articles[self.wh] = copy.deepcopy(s["excluded"])
        state.skipped_articles[self.wh] = copy.deepcopy(s.get("skipped", set()))
        state.manual_overrides[self.wh] = copy.deepcopy(s["overrides"])
        state.custom_turnovers[self.wh] = copy.deepcopy(s["turnovers"])
        
        calculate_all_warehouses()
        
        self.is_populating = True
        self.refresh_table()
        self.update_undo_redo_buttons()

    def on_add_to_export_changed(self, cb_state):
        is_checked = (cb_state == Qt.CheckState.Checked.value)
        if is_checked:
            state.export_warehouses.add(self.wh)
        else:
            state.export_warehouses.discard(self.wh)
            
        calculate_all_warehouses()
        self.refresh_table()

    def add_row(self):
        from PyQt6.QtWidgets import QInputDialog
        art, ok1 = QInputDialog.getText(self, TEXTS["add_row_title"], TEXTS["add_row_prompt1"])
        if ok1 and art.strip():
            art = art.strip()
            qty, ok2 = QInputDialog.getInt(self, TEXTS["add_row_title"], TEXTS["add_row_prompt2"], min=1, max=100000)
            if ok2:
                if self.wh not in state.custom_rows:
                    state.custom_rows[self.wh] = []
                state.custom_rows[self.wh].append({"supplierArticle": art, "quantity": qty})
                
                # If it was excluded, unexclude it
                if self.wh in state.excluded_articles and art in state.excluded_articles[self.wh]:
                    state.excluded_articles[self.wh].remove(art)
                    
                # If it already existed naturally, override its shipment
                if self.wh not in state.manual_overrides:
                    state.manual_overrides[self.wh] = {}
                state.manual_overrides[self.wh][art] = qty

                calculate_all_warehouses()
                self.save_state()
                self.refresh_table()

    def rem_row(self):
        ranges = self.table.selectedRanges()
        if not ranges:
            return
            
        arts_to_remove = set()
        art_col_idx = -1
        for c_idx, (c_key, _) in enumerate(self.current_cols):
            if c_key == "supplierArticle":
                art_col_idx = c_idx
                break
                
        if art_col_idx == -1:
            return
            
        for r in ranges:
            for row in range(r.topRow(), r.bottomRow() + 1):
                item = self.table.item(row, art_col_idx)
                if item:
                    arts_to_remove.add(item.text())
                    
        if not arts_to_remove:
            return

        if self.wh not in state.excluded_articles:
            state.excluded_articles[self.wh] = set()
            
        for a in arts_to_remove:
            state.excluded_articles[self.wh].add(a)

        calculate_all_warehouses()
        self.save_state()
        self.refresh_table()

    def setup_table(self):
        # Save header state before rebuilding
        if self.table.columnCount() > 0:
            state.set_setting("tableHeaderState", self.table.horizontalHeader().saveState())

        cols = []
        if state.visible_columns.get("subject", True): cols.append(("itemSubject", TEXTS["col_subject"]))
        if state.visible_columns.get("name", True): cols.append(("itemName", TEXTS["col_name"]))
        cols.append(("supplierArticle", TEXTS["col_article"]))
        if state.visible_columns.get("stock", True): cols.append(("current_stock", TEXTS["col_stock"]))
        if state.visible_columns.get("missing_days", True): cols.append(("missing_days", "Без остатка(дн)"))
        if state.visible_columns.get("sales", True): cols.append(("sales_30days", TEXTS["col_sales"]))
        if state.visible_columns.get("avg", True): cols.append(("avg_daily_sales", TEXTS["col_avg"]))
        if state.visible_columns.get("wb_turnover", True): cols.append(("wb_turnover", TEXTS["col_wb_turnover"]))
        if state.visible_columns.get("turnover", True): cols.append(("turnover_days", TEXTS["col_turnover"]))
        if state.visible_columns.get("target", True): cols.append(("target_stock", TEXTS["col_target"]))
        if state.visible_columns.get("need", True): cols.append(("need", TEXTS["col_need"]))
        if state.visible_columns.get("ms_stock", True): cols.append(("ms_stock", TEXTS["col_ms_stock"]))
        if state.visible_columns.get("final", True): cols.append(("final_shipment", TEXTS["col_final"]))
        
        self.current_cols = cols
        self.table.setColumnCount(len(cols))
        self.table.setHorizontalHeaderLabels([c[1] for c in cols])
        
        # Reset resize mode to stretch for equal width
        for i in range(len(cols)):
            self.table.horizontalHeader().setSectionResizeMode(i, QHeaderView.ResizeMode.Stretch)

        self.refresh_table()

    def refresh_table(self):
        self.is_populating = True
        
        data = state.calculated_data.get(self.wh, [])
        query = self.search_input.text().lower()
        
        if query:
            data = [d for d in data if any(query in str(v).lower() for v in d.values())]
            
        if getattr(state, "hide_zero_ms", False):
            data = [d for d in data if int(state.ms_stocks.get(str(d.get("supplierArticle", "")), 0)) > 0]
            
        if state.hide_zero:
            data = [d for d in data if d.get("need", 0) > 0 or d.get("final_shipment", 0) > 0]

        if state.turnover_filter_enabled:
            lo = state.turnover_filter_from
            hi = state.turnover_filter_to
            data = [d for d in data if lo <= d.get("wb_turnover", 0) <= hi]

        # Save scroll state
        vbar = self.table.verticalScrollBar()
        old_scroll = vbar.value()

        self.table.setSortingEnabled(False)
        self.table.clearContents()
        self.table.setRowCount(len(data))
        
        for row_idx, item in enumerate(data):
            for col_idx, (col_key, _) in enumerate(self.current_cols):
                if col_key == "ms_stock":
                    art = item.get("supplierArticle", "")
                    art_str = str(art)
                    base_ms = int(state.ms_stocks.get(art_str, 0))
                    
                    reserved_global = state.ms_reserved_local.get(art_str, 0)
                    
                    remain_ms = base_ms - reserved_global
                    
                    val_str = f"{base_ms} ({remain_ms})"
                else:
                    val = item.get(col_key, "")
                
                if col_key == "final_shipment":
                    cell_item = NumericTableWidgetItem()
                    cell_item.setData(Qt.ItemDataRole.UserRole, int(val) if isinstance(val, int) else val)
                    cell_item.setData(Qt.ItemDataRole.DisplayRole, "")
                else:
                    cell_item = QTableWidgetItem()
                    if col_key == "ms_stock":
                        cell_item.setData(Qt.ItemDataRole.EditRole, val_str)
                    elif isinstance(val, (int, float)):
                        cell_item.setData(Qt.ItemDataRole.EditRole, int(val) if isinstance(val, int) else val)
                    else:
                        cell_item.setText(str(val))
                
                # Alignment logic
                if col_key in ("itemName", "itemSubject"):
                    cell_item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
                else:
                    cell_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                
                if col_key == "turnover_days":
                    cell_item.setFlags(cell_item.flags() | Qt.ItemFlag.ItemIsEditable)
                else:
                    cell_item.setFlags(cell_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    
                self.table.setItem(row_idx, col_idx, cell_item)
                
                if col_key == "final_shipment":
                    is_skipped = item.get("is_skipped", False)
                    widget = FinalShipmentWidget(val, is_skipped)
                    art_code = item.get("supplierArticle")
                    widget.valueChanged.connect(lambda v, a=art_code: self.on_final_shipment_changed(a, v))
                    widget.skipChanged.connect(lambda s, a=art_code: self.on_final_shipment_skipped(a, s))
                    self.table.setCellWidget(row_idx, col_idx, widget)

                    # Green text for items added via RecentSuppliesDialog
                    if str(art_code) in state.recently_added.get(self.wh, set()):
                        widget.spin.setStyleSheet(
                            "QSpinBox { background: transparent; border: none; font-weight: bold; color: #4caf50; }"
                            "QSpinBox::up-button, QSpinBox::down-button { width: 0px; }"
                        )
                        widget.spin.setToolTip("✨ Добавлен из недавних поступлений")
                    
                if col_key == "ms_stock" and remain_ms < 0:
                    lbl = QLabel(val_str)
                    lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
                    lbl.setStyleSheet("color: #ff5252; font-weight: bold; background-color: #521520;")
                    lbl.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
                    self.table.setCellWidget(row_idx, col_idx, lbl)

            # Highlight row if item was recently received (checkbox-driven)
            art = str(item.get("supplierArticle", ""))
            if state.highlight_recent and state.recent_supplies:
                date_str = state.recent_supplies.get(art) or state.recent_supplies.get(art.lower())
                if date_str and _is_within_days(date_str, state.recent_days):
                    if " " in date_str:
                        date_str = date_str.split(" ")[0]
                    label = f"📦 Принято: {date_str}"
                    for col_idx in range(len(self.current_cols)):
                        cell = self.table.item(row_idx, col_idx)
                        if cell:
                            cell.setData(Qt.ItemDataRole.BackgroundRole, QBrush(QColor("#1a3a2a")))
                            cell.setToolTip(label)
                    
        # Restore scroll state
        self.table.verticalScrollBar().setValue(old_scroll)
        self.is_populating = False
        


    def on_item_changed(self, item):
        if self.is_populating:
            return
            
        col_idx = item.column()
        row_idx = item.row()
        col_key = self.current_cols[col_idx][0]
        
        if col_key not in ("turnover_days", "final_shipment"):
            return
            
        art_idx = next(i for i, c in enumerate(self.current_cols) if c[0] == "supplierArticle")
        art = self.table.item(row_idx, art_idx).text()
        
        try:
            val = int(item.text())
        except ValueError:
            return 
            
        if col_key == "turnover_days":
            state.custom_turnovers[self.wh][art] = val
        elif col_key == "final_shipment":
            if self.wh not in state.manual_overrides:
                state.manual_overrides[self.wh] = {}
            state.manual_overrides[self.wh][art] = val
            
        calculate_all_warehouses()
        
        updated_item = next((i for i in state.calculated_data[self.wh] if i["supplierArticle"] == str(art)), None)
        if updated_item:
            self.is_populating = True
            for c_idx, (c_key, _) in enumerate(self.current_cols):
                if c_key == "ms_stock":
                    art_str = str(art)
                    base_ms = int(state.ms_stocks.get(art_str, 0))
                    reserved_global = state.ms_reserved_local.get(art_str, 0)
                    remain_ms = base_ms - reserved_global
                    val_str = f"{base_ms} ({remain_ms})"
                    
                    if remain_ms < 0:
                        lbl = QLabel(val_str)
                        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
                        lbl.setStyleSheet("color: #ff5252; font-weight: bold; background-color: #521520;")
                        lbl.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
                        self.table.setCellWidget(row_idx, c_idx, lbl)
                    else:
                        self.table.removeCellWidget(row_idx, c_idx)
                        cell_item = self.table.item(row_idx, c_idx)
                        if not cell_item:
                            cell_item = QTableWidgetItem()
                            cell_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                            cell_item.setFlags(cell_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                            self.table.setItem(row_idx, c_idx, cell_item)
                        cell_item.setData(Qt.ItemDataRole.EditRole, val_str)
                        
                elif c_key in ("target_stock", "need"):
                    cell_item = self.table.item(row_idx, c_idx)
                    if cell_item:
                        val_item = updated_item.get(c_key, "")
                        cell_item.setData(Qt.ItemDataRole.EditRole, int(val_item) if isinstance(val_item, int) else val_item)
                        
                elif c_key == "final_shipment" and col_key != "final_shipment":
                    cell_item = self.table.item(row_idx, c_idx)
                    if cell_item:
                        val_curr = updated_item.get("final_shipment", 0)
                        cell_item.setData(Qt.ItemDataRole.EditRole, val_curr)
                    widget = self.table.cellWidget(row_idx, c_idx)
                    if widget and hasattr(widget, 'spin'):
                        val_curr = updated_item.get("final_shipment", 0)
                        if widget.spin.value() != val_curr:
                            widget.spin.blockSignals(True)
                            widget.spin.setValue(val_curr)
                            widget.spin.blockSignals(False)

            self.is_populating = False
            
        self.save_state()

    def on_final_shipment_changed(self, art, val):
        if self.is_populating: return
        if self.wh not in state.manual_overrides:
            state.manual_overrides[self.wh] = {}
        state.manual_overrides[self.wh][art] = val
        calculate_all_warehouses()
        
        updated_item = next((i for i in state.calculated_data[self.wh] if i["supplierArticle"] == str(art)), None)
        if updated_item:
            row_idx = -1
            art_col = next((ci for ci, c in enumerate(self.current_cols) if c[0] == "supplierArticle"), -1)
            for i in range(self.table.rowCount()):
                if self.table.item(i, art_col).text() == str(art):
                    row_idx = i
                    break
            
            if row_idx >= 0:
                self.is_populating = True
                for c_idx, (c_key, _) in enumerate(self.current_cols):
                    if c_key == "ms_stock":
                        art_str = str(art)
                        base_ms = int(state.ms_stocks.get(art_str, 0))
                        reserved_global = state.ms_reserved_local.get(art_str, 0)
                        remain_ms = base_ms - reserved_global
                        val_str = f"{base_ms} ({remain_ms})"
                        
                        if remain_ms < 0:
                            lbl = QLabel(val_str)
                            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
                            lbl.setStyleSheet("color: #ff5252; font-weight: bold; background-color: #521520;")
                            lbl.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
                            self.table.setCellWidget(row_idx, c_idx, lbl)
                        else:
                            self.table.removeCellWidget(row_idx, c_idx)
                            cell_item = self.table.item(row_idx, c_idx)
                            if not cell_item:
                                cell_item = QTableWidgetItem()
                                cell_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                                cell_item.setFlags(cell_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                                self.table.setItem(row_idx, c_idx, cell_item)
                            cell_item.setData(Qt.ItemDataRole.EditRole, val_str)
                            
                    elif c_key in ("target_stock", "need"):
                        cell_item = self.table.item(row_idx, c_idx)
                        if cell_item:
                            val_item = updated_item.get(c_key, "")
                            cell_item.setData(Qt.ItemDataRole.EditRole, int(val_item) if isinstance(val_item, int) else val_item)
                            
                    elif c_key == "final_shipment":
                        val_curr = updated_item.get("final_shipment", 0)
                        widget = self.table.cellWidget(row_idx, c_idx)
                        if widget and isinstance(widget, FinalShipmentWidget):
                            if widget.spin.value() != val_curr:
                                widget.spin.blockSignals(True)
                                widget.spin.setValue(val_curr)
                                widget.spin.blockSignals(False)

                self.is_populating = False
        
        self.save_state()

    def on_final_shipment_skipped(self, art, is_skipped):
        if self.is_populating: return
        if self.wh not in state.skipped_articles:
            state.skipped_articles[self.wh] = set()
        
        if is_skipped:
            state.skipped_articles[self.wh].add(art)
        else:
            state.skipped_articles[self.wh].discard(art)
            
        calculate_all_warehouses()
        
        # update inline to avoid losing focus/scroll
        updated_item = next((i for i in state.calculated_data[self.wh] if i["supplierArticle"] == str(art)), None)
        if updated_item:
            row_idx = -1
            art_col = next((ci for ci, c in enumerate(self.current_cols) if c[0] == "supplierArticle"), -1)
            for i in range(self.table.rowCount()):
                if self.table.item(i, art_col).text() == str(art):
                    row_idx = i
                    break
            if row_idx >= 0:
                self.is_populating = True
                for c_idx, (c_key, _) in enumerate(self.current_cols):
                    if c_key == "ms_stock":
                        art_str = str(art)
                        base_ms = int(state.ms_stocks.get(art_str, 0))
                        reserved_global = state.ms_reserved_local.get(art_str, 0)
                        remain_ms = base_ms - reserved_global
                        val_str = f"{base_ms} ({remain_ms})"
                        if remain_ms < 0:
                            lbl = QLabel(val_str)
                            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
                            lbl.setStyleSheet("color: #ff5252; font-weight: bold; background-color: #521520;")
                            lbl.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
                            self.table.setCellWidget(row_idx, c_idx, lbl)
                        else:
                            self.table.removeCellWidget(row_idx, c_idx)
                            cell_item = self.table.item(row_idx, c_idx)
                            if not cell_item:
                                cell_item = QTableWidgetItem()
                                cell_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                                cell_item.setFlags(cell_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                                self.table.setItem(row_idx, c_idx, cell_item)
                            cell_item.setData(Qt.ItemDataRole.EditRole, val_str)
                self.is_populating = False

        self.save_state()

    def on_header_context_menu(self, pos):
        col = self.table.horizontalHeader().logicalIndexAt(pos)
        if col >= 0:
            current_order = self.table.horizontalHeader().sortIndicatorOrder()
            new_order = Qt.SortOrder.AscendingOrder if current_order == Qt.SortOrder.DescendingOrder else Qt.SortOrder.DescendingOrder
            self.table.setSortingEnabled(True)
            self.table.sortByColumn(col, new_order)
            self.table.setSortingEnabled(False)



class RecentSuppliesDialog(QDialog):
    def __init__(self, warehouse_name, parent=None):
        super().__init__(parent)
        self.wh = warehouse_name
        self.setWindowTitle(TEXTS.get("recent_dialog_title", "Недавние поступления"))
        self.resize(750, 550)

        layout = QVBoxLayout()

        # Filter row: "за последние [N] дн."
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel(TEXTS.get("recent_filter_label", "За последние")))
        self.days_spin = QSpinBox()
        self.days_spin.setRange(1, 60)
        self.days_spin.setValue(20)
        self.days_spin.setMinimumWidth(60)
        self.days_spin.valueChanged.connect(self.apply_filter)
        filter_layout.addWidget(self.days_spin)
        filter_layout.addWidget(QLabel(TEXTS.get("recent_filter_suffix", "дн.")))
        filter_layout.addStretch()

        self.info_label = QLabel("")
        filter_layout.addWidget(self.info_label)
        layout.addLayout(filter_layout)

        # Table: Наименование | ШК | Артикул | На складе МС | К отгрузке | Добавить
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels([
            TEXTS.get("recent_col_name", "Наименование"),
            TEXTS.get("recent_col_barcode", "ШК"),
            TEXTS.get("recent_col_article", "Артикул"),
            TEXTS.get("recent_col_stock", "На складе МС"),
            TEXTS.get("recent_col_qty", "К отгрузке"),
            ""
        ])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for i in range(1, 5):
            header.setSectionResizeMode(i, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(5, 90)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)

        layout.addWidget(self.table)

        # Close button
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_close = QPushButton(TEXTS.get("btn_close", "Закрыть"))
        btn_close.clicked.connect(self.accept)
        btn_layout.addWidget(btn_close)
        layout.addLayout(btn_layout)

        self.setLayout(layout)
        self.apply_filter()

    def apply_filter(self):
        days = self.days_spin.value()
        from datetime import datetime, timedelta
        cutoff = datetime.now() - timedelta(days=days)

        # Filter supplies by date
        filtered = []
        for s in state.recent_supplies_detailed:
            try:
                d = datetime.strptime(s["received_date"][:19], "%Y-%m-%d %H:%M:%S")
            except (ValueError, TypeError):
                continue
            if d < cutoff:
                continue
            filtered.append(s)

        # Sort by date descending
        filtered.sort(key=lambda s: s["received_date"], reverse=True)

        # Separate: already in table vs not
        current_arts = set()
        data = state.calculated_data.get(self.wh, [])
        for item in data:
            current_arts.add(str(item.get("supplierArticle", "")))

        excluded = state.excluded_articles.get(self.wh, set())
        not_in_table = []
        for s in filtered:
            if s["article"] in current_arts:
                continue
            if s["article"] in excluded:
                continue
            not_in_table.append(s)

        self.table.setRowCount(0)
        self.table.setRowCount(len(not_in_table))
        self.table.clearContents()

        for row_idx, s in enumerate(not_in_table):
            # Name
            name_item = QTableWidgetItem(s["name"])
            name_item.setFlags(name_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            name_item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
            self.table.setItem(row_idx, 0, name_item)

            # Barcode
            bc_item = QTableWidgetItem(s["barcode"])
            bc_item.setFlags(bc_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            bc_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(row_idx, 1, bc_item)

            # Article
            art_item = QTableWidgetItem(s["article"])
            art_item.setFlags(art_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            art_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(row_idx, 2, art_item)

            # MS stock
            stock = int(state.ms_stocks.get(s["article"], 0))
            stock_item = QTableWidgetItem(str(stock))
            stock_item.setFlags(stock_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            stock_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            if stock == 0:
                stock_item.setForeground(QColor("#ff5252"))
            self.table.setItem(row_idx, 3, stock_item)

            # Spinbox for quantity
            spin = QSpinBox()
            spin.setRange(1, 100000)
            spin.setValue(1)
            spin.setMinimumWidth(60)
            self.table.setCellWidget(row_idx, 4, spin)

            # Add button
            btn = QPushButton(TEXTS.get("recent_btn_add", "Добавить"))
            btn.clicked.connect(lambda checked, art=s["article"], sp=spin: self.add_to_table(art, sp.value()))
            self.table.setCellWidget(row_idx, 5, btn)

        total = len(not_in_table)
        self.info_label.setText(TEXTS.get("recent_info", "Показано: {count}").format(count=total))

    def add_to_table(self, article, qty):
        if self.wh not in state.custom_rows:
            state.custom_rows[self.wh] = []
        state.custom_rows[self.wh].append({"supplierArticle": article, "quantity": qty})

        if self.wh not in state.manual_overrides:
            state.manual_overrides[self.wh] = {}
        state.manual_overrides[self.wh][article] = qty

        # Track for green highlighting
        if self.wh not in state.recently_added:
            state.recently_added[self.wh] = set()
        state.recently_added[self.wh].add(article)

        # Remove from excluded if it was there
        if self.wh in state.excluded_articles and article in state.excluded_articles[self.wh]:
            state.excluded_articles[self.wh].discard(article)

        calculate_all_warehouses()

        # Refresh the main warehouse tab
        parent = self.parent()
        if parent and hasattr(parent, "refresh"):
            parent.refresh()

        # Re-filter (remove from list)
        self.apply_filter()

        QMessageBox.information(self, TEXTS.get("msg_added_title", "Добавлено"),
            TEXTS.get("msg_added_body", "Товар {art} ({qty} шт.) добавлен в таблицу.").format(art=article, qty=qty))


class WarehouseSettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(TEXTS["wh_settings_title"])
        self.resize(500, 600)
        
        layout = QVBoxLayout()
        
        self.btn_hide_empty = QCheckBox(TEXTS["wh_settings_hide"])
        self.btn_hide_empty.setChecked(state.auto_hide_empty)
        self.btn_hide_empty.stateChanged.connect(self.on_auto_hide_changed)
        layout.addWidget(self.btn_hide_empty)
        
        self.use_missing_days_cb = QCheckBox(TEXTS.get("cb_use_missing_days", "Учитывать дни без остатка в оборачиваемости"))
        self.use_missing_days_cb.setChecked(getattr(state, "use_missing_days", False))
        self.use_missing_days_cb.stateChanged.connect(self.on_use_missing_days_changed)
        layout.addWidget(self.use_missing_days_cb)
        
        self.list_widget = QListWidget()
        self.list_widget.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        
        self.list_items = {}
        for w in state.warehouses:
            item = QListWidgetItem(w)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked if w in state.active_warehouses else Qt.CheckState.Unchecked)
            self.list_items[w] = item
            self.list_widget.addItem(item)
            
        layout.addWidget(self.list_widget)
        
        # Themes dropdown
        theme_layout = QHBoxLayout()
        theme_layout.addWidget(QLabel("Цветовая тема:"))
        self.theme_combo = QComboBox()
        themes_list = list_themes()
        self.theme_combo.addItems(themes_list)
        current_theme = getattr(state, "theme", "dark_teal.xml")
        if current_theme in themes_list:
            self.theme_combo.setCurrentText(current_theme)
        self.theme_combo.currentTextChanged.connect(self.on_theme_changed)
        theme_layout.addWidget(self.theme_combo)
        layout.addLayout(theme_layout)

        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)
        
        self.setLayout(layout)
        
        
    def on_theme_changed(self, theme_name):
        state.set_setting("theme", theme_name)
        state.theme = theme_name
        app = QApplication.instance()
        apply_stylesheet(app, theme=theme_name)
        
        # Override selection background color for text inputs globally (same as main)
        extra_style = """
        QLineEdit {
            selection-background-color: #555555;
            selection-color: #ffffff;
        }
        QAbstractSpinBox {
            selection-background-color: #555555;
            selection-color: #ffffff;
        }
        QTableWidget {
            alternate-background-color: rgba(128, 128, 128, 0.1);
        }
        """
        app.setStyleSheet(app.styleSheet() + extra_style)

    def on_auto_hide_changed(self, state_val):
        is_checked = state_val == Qt.CheckState.Checked.value
        state.set_setting("auto_hide_empty", is_checked)
        state.auto_hide_empty = is_checked
        if is_checked:
            self.hide_empty()
            
    def hide_empty(self):
        for w in state.warehouses:
            w_stocks = [s for s in state.stocks_data if s["warehouseName"] == w]
            
            total_qty = sum(s.get("quantity", 0) for s in w_stocks)
            
            if total_qty == 0:
                if w in self.list_items:
                    self.list_items[w].setCheckState(Qt.CheckState.Unchecked)

    def on_use_missing_days_changed(self, state_val):
        is_checked = state_val == Qt.CheckState.Checked.value
        state.set_setting("use_missing_days", is_checked)
        state.use_missing_days = is_checked

class WarehousesScreen(QWidget):
    def __init__(self, parent_nav):
        super().__init__()
        self.parent_nav = parent_nav
        
        layout = QVBoxLayout()
        
        header_layout = QHBoxLayout()
        self.title_label = QLabel(TEXTS["step2_title"])
        self.title_label.setFont(QFont("Arial", 16, QFont.Weight.Bold))
        header_layout.addWidget(self.title_label)
        
        btn_update_ms = QPushButton(TEXTS.get("btn_update_ms", "🔄 Обновить МС"))
        btn_update_ms.clicked.connect(self.update_ms_stocks)
        header_layout.addWidget(btn_update_ms)
        
        btn_settings = QPushButton(TEXTS["btn_settings"])
        btn_settings.clicked.connect(self.open_settings)
        header_layout.addWidget(btn_settings)

        # Column visibility menu
        self.cols_menu_btn = QPushButton("Настройка колонок")
        self.cols_menu_btn.setMinimumWidth(150)
        self.cols_menu = QMenu(self)

        cols_def = [
            ("subject", TEXTS["col_subject"], "Предмет/категория товара"),
            ("name", TEXTS["col_name"], "Наименование товара"),
            ("stock", TEXTS["col_stock"], "Текущий остаток на нашем складе (из отчета)"),
            ("missing_days", "Без остатка(дн)", "Сколько дней товара не было на складе (за 30 дней)"),
            ("sales", "Заказы(30д)", "Количество заказов из отчета за 30 дней"),
            ("avg", "Средн. в день", "Среднее количество заказов в день"),
            ("wb_turnover", "Оборач.(Тек)", "Оборачиваемость на текущий момент"),
            ("turnover", "Оборач.(План)", "Целевое значение оборачиваемости"),
            ("target", "Цель", "Желаемый остаток на складе (Ср.в день * Оборач.(План))"),
            ("need", "План", "Количество, которое нужно поставить (Цель - Текущий остаток)"),
            ("ms_stock", "Ост.МС", "Доступные остатки в Моем Складе"),
            ("final", TEXTS["col_final"], "Итоговое количество к отгрузке.")
        ]

        for col_key, label_text, tooltip in cols_def:
            cb = QCheckBox(label_text)
            cb.setToolTip(tooltip)
            if col_key == "article":
                continue  # always visible, handled in tab
            cb.setChecked(state.visible_columns.get(col_key, True))
            cb.stateChanged.connect(lambda state_val, k=col_key: self.on_column_toggled(k, state_val))

            container = QWidget()
            cbl = QHBoxLayout(container)
            cbl.setContentsMargins(5, 2, 5, 2)
            cbl.addWidget(cb)
            wa = QWidgetAction(self)
            wa.setDefaultWidget(container)
            self.cols_menu.addAction(wa)

        self.cols_menu_btn.setMenu(self.cols_menu)
        header_layout.addWidget(self.cols_menu_btn)

        # Recent supplies button
        self.recent_btn = QPushButton(TEXTS.get("btn_recent_supplies", "Недавние поступления"))
        self.recent_btn.clicked.connect(self.open_recent_supplies)
        header_layout.addWidget(self.recent_btn)

        header_layout.addStretch()
        
        btn_next = QPushButton(TEXTS["btn_next"])
        btn_next.setProperty('class', 'success')
        btn_next.clicked.connect(self.go_to_preview)
        header_layout.addWidget(btn_next)
        
        layout.addLayout(header_layout)
        
        self.tabs = QTabWidget()
        self.tabs.setUsesScrollButtons(True)
        self.tabs.currentChanged.connect(self.on_tab_changed)
        self._populating_tabs = False
        layout.addWidget(self.tabs)
        
        self.setLayout(layout)

    def on_tab_changed(self, index):
        if self._populating_tabs:
            return
        widget = self.tabs.widget(index)
        if widget and hasattr(widget, 'refresh_table'):
            widget.refresh_table()

    def open_settings(self):
        dlg = WarehouseSettingsDialog(self)
        if dlg.exec():
            ordered_warehouses = []
            active_warehouses = []
            for i in range(dlg.list_widget.count()):
                item = dlg.list_widget.item(i)
                w = item.text()
                ordered_warehouses.append(w)
                if item.checkState() == Qt.CheckState.Checked:
                    active_warehouses.append(w)
            
            state.warehouses = ordered_warehouses
            state.active_warehouses = active_warehouses
            state.set_setting("warehouses_order", state.warehouses)
            state.set_setting("active_warehouses", state.active_warehouses)
            
            calculate_all_warehouses()
            self.refresh()

    def update_ms_stocks(self):
        try:
            self.window().statusBar().showMessage("Обновление остатков из Моего Склада...")
            QApplication.processEvents()
            state.ms_stocks_details = moysklad_api.get_all_stocks()
            state.ms_stocks = {}
            acc = getattr(state, "current_wb_account", "WB1")
            target_folder = ("магазин на wildberries" if acc == "WB1" else "магазин на wildberries юля").lower()
            for k, v in state.ms_stocks_details.items():
                folder = v.get("folder", "Без группы").strip().lower()
                if folder == target_folder:
                    state.ms_stocks[str(k)] = v["stock"]
            self.window().statusBar().showMessage("Остатки успешно обновлены!", 3000)
            
            calculate_all_warehouses()
            
            current_tab = self.tabs.currentWidget()
            if current_tab:
                current_tab.is_populating = True
                current_tab.refresh_table()
                current_tab.save_state()
        except Exception as e:
            QMessageBox.critical(self, "Ошибка обновления", f"Произошла ошибка при получении остатков:\\n{e}")

    def go_to_preview(self):
        if self.tabs.count() > 0:
            current_idx = self.tabs.currentIndex()
            state.current_warehouse = self.tabs.tabText(current_idx)
        self.parent_nav.go_to_preview()

    def refresh(self):
        # Update title based on account
        acc = getattr(state, "current_wb_account", "")
        if acc:
            self.title_label.setText(f"{acc}")
        else:
            self.title_label.setText(TEXTS["step2_title"])

        self._populating_tabs = True
        current_tabs = [self.tabs.tabText(i) for i in range(self.tabs.count())]
        if current_tabs != state.active_warehouses:
            self.tabs.clear()
            for w in state.active_warehouses:
                self.tabs.addTab(WarehouseTabWidget(w), w)
        else:
            for i in range(self.tabs.count()):
                widget = self.tabs.widget(i)
                widget.refresh_table()
        self._populating_tabs = False

    def on_column_toggled(self, key, check_state):
        is_checked = (check_state == Qt.CheckState.Checked.value)
        state.visible_columns[key] = is_checked
        state.set_setting(f"col_{key}", is_checked)
        self.refresh()

    def open_recent_supplies(self):
        current_tab = self.tabs.currentWidget()
        if not current_tab or not hasattr(current_tab, "wh"):
            return
        if not state.recent_supplies_detailed:
            QMessageBox.information(self,
                TEXTS.get("msg_no_supplies_title", "Нет данных"),
                TEXTS.get("msg_no_supplies_body", "Нет данных о поступлениях. Перезагрузите Excel-файл."))
            return
        dlg = RecentSuppliesDialog(current_tab.wh, parent=self)
        dlg.exec()


class MSShipmentDialog(QDialog):
    def __init__(self, warehouse_name, parent=None):
        super().__init__(parent)
        self.warehouse_name = warehouse_name
        self.setWindowTitle(TEXTS["ms_dialog_title"].format(warehouse_name=warehouse_name))
        self.resize(500, 250)

        layout = QVBoxLayout()

        # Mode selection: new or existing
        mode_layout = QHBoxLayout()
        self.rb_new = QRadioButton(TEXTS["ms_dialog_create_new"])
        self.rb_new.setChecked(True)
        self.rb_existing = QRadioButton(TEXTS["ms_dialog_add_existing"])
        self.rb_new.toggled.connect(self.on_mode_changed)
        mode_layout.addWidget(self.rb_new)
        mode_layout.addWidget(self.rb_existing)
        mode_layout.addStretch()
        layout.addLayout(mode_layout)

        # Existing shipment selector
        self.existing_layout = QVBoxLayout()
        lbl_select = QLabel(TEXTS["ms_dialog_select_shipment"])
        self.existing_layout.addWidget(lbl_select)
        self.cb_shipments = QComboBox()
        self.cb_shipments.setMinimumWidth(350)
        self.cb_shipments.setEnabled(False)
        self.existing_layout.addWidget(self.cb_shipments)
        self.shipment_data = []
        layout.addLayout(self.existing_layout)

        # Name and description
        form = QFormLayout()
        self.edit_name = QLineEdit()
        self.edit_name.setPlaceholderText(TEXTS["ms_dialog_ph_name"])
        self.edit_desc = QLineEdit()
        self.edit_desc.setPlaceholderText(TEXTS["ms_dialog_ph_desc"])
        if getattr(state, "current_wb_account", "WB1") == "WB2":
            self.edit_desc.setText("2 каб")

        form.addRow(TEXTS["ms_dialog_lbl_name"], self.edit_name)
        form.addRow(TEXTS["ms_dialog_lbl_desc"], self.edit_desc)
        layout.addLayout(form)

        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

        self.setLayout(layout)
        self.load_shipments()

    def on_mode_changed(self):
        is_existing = self.rb_existing.isChecked()
        self.cb_shipments.setEnabled(is_existing)
        self.edit_name.setEnabled(not is_existing)

    def load_shipments(self):
        """Load recent shipments (last 14 days) in background."""
        try:
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
            self.cb_shipments.addItem(TEXTS["ms_dialog_loading_shipments"])
            QApplication.processEvents()

            self.shipment_data = moysklad_api.get_demands(days_back=14, limit=100)

            self.cb_shipments.clear()
            if not self.shipment_data:
                self.cb_shipments.addItem("— нет отгрузок за последние 2 недели —")
                self.rb_existing.setEnabled(False)
            else:
                for d in self.shipment_data:
                    addr = d.get("shipmentAddress", "")
                    moment = d.get("moment", "")[:10] if d.get("moment") else ""
                    label = f"{d['name']} — {addr} ({moment})" if addr else f"{d['name']} ({moment})"
                    if d.get("description"):
                        label += f" [{d['description']}]"
                    self.cb_shipments.addItem(label.strip())
        except Exception:
            self.cb_shipments.clear()
            self.cb_shipments.addItem("— ошибка загрузки —")
            self.rb_existing.setEnabled(False)
        finally:
            QApplication.restoreOverrideCursor()


class DuplicatePositionDialog(QDialog):
    """Диалог разрешения конфликтов: товары уже есть в отгрузке — пропустить/заменить/добавить."""

    def __init__(self, shipment_name, duplicates, new_positions, parent=None):
        """
        duplicates: list of {art, existing_qty, existing_id, new_qty, price, meta}
        new_positions: list of {art, quantity, meta, price} — позиции без конфликтов
        """
        super().__init__(parent)
        self.duplicates = duplicates
        self.new_positions = new_positions
        self.setWindowTitle(TEXTS["dup_dialog_title"])
        self.resize(650, 400)

        layout = QVBoxLayout()

        desc = QLabel(TEXTS["dup_dialog_desc"].format(name=shipment_name))
        desc.setWordWrap(True)
        layout.addWidget(desc)

        # Table: article | existing_qty | new_qty | action | result
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels([
            TEXTS["dup_col_article"], TEXTS["dup_col_existing"],
            TEXTS["dup_col_new"], TEXTS["dup_col_action"], TEXTS["dup_col_result"]
        ])
        self.table.setRowCount(len(duplicates))
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self.table.setEditTriggers(QTableWidget.EditTrigger.AllEditTriggers)

        self.actions = []
        self.result_labels = []

        for r, d in enumerate(duplicates):
            self.table.setItem(r, 0, QTableWidgetItem(d["art"]))
            self.table.setItem(r, 1, QTableWidgetItem(str(d["existing_qty"])))

            # Editable new qty
            qty_item = QTableWidgetItem()
            qty_item.setData(Qt.ItemDataRole.EditRole, d["new_qty"])
            self.table.setItem(r, 2, qty_item)
            for col in range(3):
                self.table.item(r, col).setFlags(
                    self.table.item(r, col).flags() & ~Qt.ItemFlag.ItemIsEditable)

            # Action combo
            combo = QComboBox()
            combo.addItems([TEXTS["dup_action_skip"], TEXTS["dup_action_replace"], TEXTS["dup_action_add"]])
            combo.setCurrentIndex(2)  # default: add
            combo.currentIndexChanged.connect(lambda idx, row=r: self.update_result(row))
            self.table.setCellWidget(r, 3, combo)
            self.actions.append(combo)

            # Result label
            result = QLabel(str(d["new_qty"]))
            result.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setCellWidget(r, 4, result)
            self.result_labels.append(result)

        layout.addWidget(self.table)

        # Summary
        new_count = len(new_positions)
        new_qty = sum(p["quantity"] for p in new_positions)
        self.lbl_new = QLabel(TEXTS["dup_summary_new"].format(count=new_count, qty=new_qty))
        layout.addWidget(self.lbl_new)

        # Total summary (will update dynamically)
        self.lbl_total = QLabel()
        self.update_total()
        layout.addWidget(self.lbl_total)

        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

        self.setLayout(layout)

    def update_result(self, row):
        """Update result label based on action choice."""
        combo = self.actions[row]
        d = self.duplicates[row]
        qty_item = self.table.item(row, 2)
        try:
            new_qty = int(qty_item.data(Qt.ItemDataRole.EditRole))
        except (ValueError, TypeError):
            new_qty = d["new_qty"]

        action = combo.currentIndex()
        # 0: skip, 1: replace, 2: add
        if action == 0:
            result = "—"
        elif action == 1:
            result = str(new_qty)
        else:
            result = str(d["existing_qty"] + new_qty)

        self.result_labels[row].setText(result)
        self.update_total()

    def update_total(self):
        new_qty = sum(p["quantity"] for p in self.new_positions)
        new_count = len(self.new_positions)

        dup_total = 0
        dup_count = 0
        for r, d in enumerate(self.duplicates):
            action = self.actions[r].currentIndex()
            if action == 0:  # skip
                continue
            qty_item = self.table.item(r, 2)
            try:
                nq = int(qty_item.data(Qt.ItemDataRole.EditRole))
            except (ValueError, TypeError):
                nq = d["new_qty"]
            if action == 1:  # replace
                dup_total += nq
            else:  # add
                dup_total += d["existing_qty"] + nq
            dup_count += 1

        total_count = new_count + dup_count
        total_qty = new_qty + dup_total
        self.lbl_total.setText(TEXTS["dup_summary_total"].format(count=total_count, qty=total_qty))

    def get_result(self):
        """Returns (to_add, to_update, to_delete_ids, skipped_arts)."""
        to_add = list(self.new_positions)  # non-conflicting
        to_update = []  # list of (position_id, new_quantity, price)
        to_delete = []
        skipped = []

        for r, d in enumerate(self.duplicates):
            action = self.actions[r].currentIndex()
            qty_item = self.table.item(r, 2)
            try:
                nq = int(qty_item.data(Qt.ItemDataRole.EditRole))
            except (ValueError, TypeError):
                nq = d["new_qty"]

            if action == 0:  # skip
                skipped.append(d["art"])
            elif action == 1:  # replace: delete old, add new
                to_delete.append(d["existing_id"])
                to_add.append({
                    "meta": d["meta"],
                    "quantity": nq,
                    "price": d["price"]
                })
            else:  # add: update existing quantity
                to_update.append((d["existing_id"], d["existing_qty"] + nq, d["price"]))

        return to_add, to_update, to_delete, skipped


class PreviewScreen(QWidget):
    def __init__(self, parent_nav):
        super().__init__()
        self.parent_nav = parent_nav
        
        layout = QVBoxLayout()
        
        header_layout = QHBoxLayout()
        btn_back = QPushButton("< Назад")
        btn_back.clicked.connect(self.parent_nav.go_to_warehouses)
        header_layout.addWidget(btn_back)
        
        title = QLabel("Итоговая матрица отгрузок")
        title.setFont(QFont("Arial", 16, QFont.Weight.Bold))
        header_layout.addWidget(title)
        
        header_layout.addStretch()
        
        self.lbl_total = QLabel("")
        header_layout.addWidget(self.lbl_total)
        
        layout.addLayout(header_layout)
        
        search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText(TEXTS["search_step3"])
        self.search_input.setMaximumWidth(300)
        self.search_input.textChanged.connect(self.on_search)
        search_layout.addWidget(self.search_input)
        search_layout.addStretch()
        layout.addLayout(search_layout)
        
        self.table = CustomTableWidget()
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("""
            QTableWidget { alternate-background-color: #2b2b2b; }
            QTableWidget::item:selected {
                background-color: rgba(255, 255, 255, 0.05);
                border: 2px solid #ffca28;
                color: #ffffff;
            }
        """)
        layout.addWidget(self.table)
        
        self.setLayout(layout)

    def refresh(self):
        self.table.clearContents()
        
        active_w = sorted(list(state.export_warehouses))
        
        if not active_w:
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            self.lbl_total.setText(TEXTS["total_units"].format(grand_total="<b>0</b>"))
            return

        shipment_items = {}
        
        for w in active_w:
            for item in state.calculated_data.get(w, []):
                if item.get("is_skipped", False):
                    continue
                val = item.get("final_shipment", 0)
                if val > 0:
                    art = str(item["supplierArticle"])
                    if art not in shipment_items:
                        shipment_items[art] = {
                            "name": item.get("itemName", ""),
                            "subject": item.get("itemSubject", ""),
                            "ships": {}
                        }
                    shipment_items[art]["ships"][w] = val
                    
        sorted_arts = sorted(list(shipment_items.keys()))
        
        cols = [TEXTS.get("col_name", "Название"), TEXTS.get("col_article", "Артикул"), "Остаток МС"] + active_w + ["ИТОГО"]
        self.table.setColumnCount(len(cols))
        self.table.setHorizontalHeaderLabels(cols)
        
        self.table.setRowCount(len(sorted_arts) + 1)
        
        lbl_actions = QTableWidgetItem("ДЕЙСТВИЯ:")
        lbl_actions.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        lbl_actions.setFont(QFont("Arial", 10, QFont.Weight.Bold))
        self.table.setItem(0, 2, lbl_actions)
        
        for w_idx, w in enumerate(active_w):
            col_idx = 3 + w_idx
            btn_panel = QWidget()
            l = QHBoxLayout(btn_panel)
            l.setContentsMargins(2, 2, 2, 2)
            b_ex = QPushButton(TEXTS.get("btn_excel", "Excel"))
            b_ms = QPushButton(TEXTS.get("btn_ms", "В МС"))
            
            b_ex.clicked.connect(lambda checked, wh=w: self.export_warehouse(wh))
            b_ms.clicked.connect(lambda checked, wh=w: self.ms_warehouse(wh))
            
            l.addWidget(b_ex)
            l.addWidget(b_ms)
            self.table.setCellWidget(0, col_idx, btn_panel)
            self.table.horizontalHeader().setSectionResizeMode(col_idx, QHeaderView.ResizeMode.ResizeToContents)

        grand_total = 0
        
        for row_offset, art in enumerate(sorted_arts):
            r = row_offset + 1
            data = shipment_items[art]
            base_ms = int(state.ms_stocks.get(art, 0))
            
            self.table.setItem(r, 0, QTableWidgetItem(data["name"]))
            self.table.setItem(r, 1, QTableWidgetItem(art))
            
            ms_item = QTableWidgetItem(str(base_ms))
            ms_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(r, 2, ms_item)
            
            accumulated = 0
            row_total = 0
            
            for w_idx, w in enumerate(active_w):
                col_idx = 3 + w_idx
                val = data["ships"].get(w, 0)
                
                accumulated += val
                row_total += val
                remain = base_ms - accumulated
                
                if remain < 0 and val > 0:
                    lbl = QLabel(str(val))
                    lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
                    lbl.setStyleSheet("background-color: #8b0000; color: #ffffff; font-weight: bold;")
                    self.table.setCellWidget(r, col_idx, lbl)
                else:
                    cell = QTableWidgetItem(str(val) if val > 0 else "")
                    cell.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    self.table.setItem(r, col_idx, cell)
                
            grand_total += row_total
            t_cell = QTableWidgetItem()
            t_cell.setData(Qt.ItemDataRole.EditRole, row_total)
            t_cell.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            t_cell.setFont(QFont("Arial", 10, QFont.Weight.Bold))
            self.table.setItem(r, len(cols) - 1, t_cell)
            
        self.lbl_total.setText(TEXTS["total_units"].format(grand_total=f"<b>{grand_total}</b>"))
        self.table.resizeColumnsToContents()
        self.on_search() # Reapply search if active

    def on_search(self):
        query = self.search_input.text().lower()
        for r in range(1, self.table.rowCount()):
            name_item = self.table.item(r, 0)
            art_item = self.table.item(r, 1)
            if not name_item or not art_item:
                continue
            name_txt = name_item.text().lower()
            art_txt = art_item.text().lower()
            if query in name_txt or query in art_txt:
                self.table.setRowHidden(r, False)
            else:
                self.table.setRowHidden(r, True)

    def export_warehouse(self, w):
        path, _ = QFileDialog.getSaveFileName(self, f"Сохранить отгрузку {w}", f"отгрузка_{w}.csv".replace(' ', '_'), "CSV Files (*.csv)")
        if not path:
            return
            
        rows = []
        for item in state.calculated_data.get(w, []):
            if item.get("is_skipped", False):
                continue
            if item["final_shipment"] > 0:
                rows.append({
                    "Артикул WB": item["supplierArticle"],
                    "Баркод": item.get("barcode", ""),
                    "Количество": item["final_shipment"]
                })
        
        if not rows:
            QMessageBox.information(self, TEXTS["msg_empty_title"], TEXTS["msg_empty_export"].format(w=w))
            return
            
        import pandas as pd
        df = pd.DataFrame(rows)
        df.to_csv(path, index=False, encoding="utf-8-sig")
        QMessageBox.information(self, TEXTS["msg_success_title"], TEXTS["msg_success_export"].format(w=w))

    def ms_warehouse(self, w):
        positions = []
        # Keep a parallel map: art -> position for duplicate detection
        pos_by_art = {}
        for item in state.calculated_data.get(w, []):
            if item.get("is_skipped", False):
                continue
            qty = item.get("final_shipment", 0)
            if qty > 0:
                art = str(item["supplierArticle"])
                details = state.ms_stocks_details.get(art)
                if not details or not details.get("meta"):
                    continue
                pos = {
                    "meta": details["meta"],
                    "quantity": qty,
                    "price": details.get("price", 0)
                }
                positions.append(pos)
                pos_by_art[art] = pos

        if not positions:
            QMessageBox.warning(self, TEXTS["msg_empty_title"], TEXTS["msg_empty_ms"].format(w=w))
            return

        dlg = MSShipmentDialog(w, self)
        if dlg.exec():
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
            try:
                if dlg.rb_existing.isChecked() and dlg.shipment_data:
                    idx = dlg.cb_shipments.currentIndex()
                    if idx < 0 or idx >= len(dlg.shipment_data):
                        QApplication.restoreOverrideCursor()
                        QMessageBox.warning(self, "Ошибка", "Не выбрана существующая отгрузка.")
                        return

                    demand = dlg.shipment_data[idx]
                    demand_href = demand["meta"]["href"]

                    # Fetch existing positions to detect duplicates
                    existing = moysklad_api.get_demand_positions(demand_href)

                    # Index existing by assortment href for comparison
                    existing_by_href = {}
                    for ep in existing:
                        href = ep["meta"].get("href", "") if ep["meta"] else ""
                        if href:
                            existing_by_href[href] = ep

                    # Find duplicates and new positions
                    dups = []
                    fresh = []
                    for pos in positions:
                        href = pos["meta"].get("href", "") if isinstance(pos["meta"], dict) else ""
                        if href and href in existing_by_href:
                            ep = existing_by_href[href]
                            # Find the article for this position
                            art = ""
                            for a, p in pos_by_art.items():
                                if p is pos:
                                    art = a
                                    break
                            dups.append({
                                "art": art,
                                "existing_qty": int(ep["quantity"]),
                                "existing_id": ep["id"],
                                "new_qty": int(pos["quantity"]),
                                "price": pos.get("price", 0),
                                "meta": pos["meta"],
                            })
                        else:
                            fresh.append(pos)

                    # Show duplicate dialog if needed
                    if dups:
                        QApplication.restoreOverrideCursor()
                        dup_dlg = DuplicatePositionDialog(demand["name"], dups, fresh, self)
                        if not dup_dlg.exec():
                            return  # user cancelled
                        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)

                        to_add, to_update, to_delete, skipped = dup_dlg.get_result()

                        # Delete replaced items first
                        for pos_id in to_delete:
                            moysklad_api.delete_position(demand_href, pos_id)

                        # Update quantities
                        for pos_id, qty, price in to_update:
                            moysklad_api.update_position(demand_href, pos_id, qty, price)

                        # Add new positions
                        if to_add:
                            moysklad_api.add_positions_to_demand(demand_href, to_add)

                        QApplication.restoreOverrideCursor()
                        QMessageBox.information(self, TEXTS["msg_success_title"],
                            TEXTS["msg_success_ms_add"].format(name=demand["name"]))
                    else:
                        # No duplicates, just add
                        moysklad_api.add_positions_to_demand(demand_href, positions)
                        QApplication.restoreOverrideCursor()
                        QMessageBox.information(self, TEXTS["msg_success_title"],
                            TEXTS["msg_success_ms_add"].format(name=demand["name"]))
                else:
                    doc_name = dlg.edit_name.text().strip()
                    description = dlg.edit_desc.text().strip()
                    org_meta = moysklad_api.get_organization("KeshFix")
                    store_meta = moysklad_api.get_store("основной")
                    cp_meta = moysklad_api.get_counterparty_meta()
                    state_meta = moysklad_api.get_state("demand", "Создаётся поставка")

                    resp = moysklad_api.create_demand_v2(
                        positions, org_meta, store_meta, cp_meta, state_meta, doc_name, description, shipment_address=w
                    )
                    QApplication.restoreOverrideCursor()
                    QMessageBox.information(self, TEXTS["msg_success_title"], TEXTS["msg_success_ms"].format(name=resp.get("name", "")))
            except Exception as ex:
                QApplication.restoreOverrideCursor()
                QMessageBox.critical(self, TEXTS["error_api_title"], TEXTS["msg_error_ms"].format(ex=str(ex)))

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Создание отгрузок WB -> МойСклад")
        self.resize(1200, 800)
        
        self.stack = QStackedWidget()
        
        self.screen_start = StartScreen(self)
        self.screen_warehouses = WarehousesScreen(self)
        self.screen_preview = PreviewScreen(self)
        
        self.stack.addWidget(self.screen_start)
        self.stack.addWidget(self.screen_warehouses)
        self.stack.addWidget(self.screen_preview)
        
        self.setCentralWidget(self.stack)

    def go_to_start(self):
        self.stack.setCurrentWidget(self.screen_start)
        
    def go_to_warehouses(self):
        self.screen_warehouses.refresh()
        self.stack.setCurrentWidget(self.screen_warehouses)
        
    def go_to_preview(self):
        self.screen_preview.refresh()
        self.stack.setCurrentWidget(self.screen_preview)

    def mousePressEvent(self, event):
        # Deselect and unfocus when empty space is clicked globally
        focused = self.focusWidget()
        if focused and hasattr(focused, "clearFocus"):
            focused.clearFocus()
        for table in self.findChildren(QTableWidget):
            table.clearSelection()
        super().mousePressEvent(event)


def main():
    app = QApplication(sys.argv)
    
    current_theme = getattr(state, "theme", "dark_teal.xml")
    apply_stylesheet(app, theme=current_theme) 
    
    # Override selection background color for text inputs globally
    extra_style = """
    QLineEdit {
        selection-background-color: #555555;
        selection-color: #ffffff;
    }
    QAbstractSpinBox {
        selection-background-color: #555555;
        selection-color: #ffffff;
    }
    QTableWidget {
        alternate-background-color: rgba(128, 128, 128, 0.1);
    }
    """
    app.setStyleSheet(app.styleSheet() + extra_style)
    
    window = MainWindow()
    window.show()
    
    # Save window geometry if we wanted, but not requested currently.
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
