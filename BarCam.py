import sys
import cv2
import time
import os
import serial
import numpy as np
import serial.tools.list_ports
import platform
import threading
import pandas as pd
import json
from PyQt5.QtWidgets import (
    QApplication, QLabel, QPushButton, QVBoxLayout, QHBoxLayout, QWidget, QFileDialog,
    QComboBox, QMessageBox, QMainWindow, QLineEdit, QTabWidget, QCheckBox, QTableWidget,
    QTableWidgetItem, QHeaderView
)
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtGui import QImage, QPixmap
from pyzbar import pyzbar
from pyzbar.pyzbar import ZBarSymbol
from pylibdmtx.pylibdmtx import decode as dmtx_decode
class BarcodeApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Barcode Inspector")
        self.setGeometry(100, 100, 1300, 800)

        self.settings_file = "settings.json"
        self.load_settings()

        # Core variables
        self.capture = None
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_frame)
        self.camera_index = None
        self.last_barcode = None
        self.barcode_set = set()  # Store scanned SKUs
        self.save_location = ""
        self.serial_port = None

        # Button colors
        self.button_colors = {
            "start": "#4CAF50",       # Green
            "stop": "#F44336",        # Red
            "snapshot": "#2196F3",    # Blue
            "clear": "#FF9800",       # Orange
            "select_folder": "#9C27B0", # Purple
            "export": "#00BCD4"       # Cyan
        }

        self.init_ui()
        self.populate_serial_ports()
        self.populate_camera_indices()
        self.set_stylesheet()

    # ------------------------ UI Setup ------------------------
    def init_ui(self):
        self.tabs = QTabWidget()

        # --- Scanner Tab ---
        self.image_label = QLabel("Camera Feed")
        self.image_label.setAlignment(Qt.AlignCenter)
        self.image_label.setMinimumSize(640, 480)

        self.barcode_label = QLabel("Current Barcode: None")
        self.count_label = QLabel("Unique SKUs Scanned: 0")

        self.serial_port_combo = QComboBox()
        self.camera_combo = QComboBox()
        self.order_input = QLineEdit()
        self.order_input.setPlaceholderText("Enter Order Number")

        # Buttons
        self.start_button = QPushButton("Start Camera")
        self.stop_button = QPushButton("Stop Camera")
        self.snapshot_button = QPushButton("Capture Snapshot")
        self.clear_button = QPushButton("Clear Session")
        self.select_folder_button = QPushButton("Select Save Folder")
        self.export_button = QPushButton("Export to Excel")

        # Apply colors
        self.start_button.setStyleSheet(f"background-color: {self.button_colors['start']}; color: white;")
        self.stop_button.setStyleSheet(f"background-color: {self.button_colors['stop']}; color: white;")
        self.snapshot_button.setStyleSheet(f"background-color: {self.button_colors['snapshot']}; color: white;")
        self.clear_button.setStyleSheet(f"background-color: {self.button_colors['clear']}; color: white;")
        self.select_folder_button.setStyleSheet(f"background-color: {self.button_colors['select_folder']}; color: white;")
        self.export_button.setStyleSheet(f"background-color: {self.button_colors['export']}; color: white;")

        self.status_label = QLabel("Status: Not Connected")

        # Table for scanned SKUs
        self.sku_table = QTableWidget()
        self.sku_table.setColumnCount(2)
        self.sku_table.setHorizontalHeaderLabels(["Timestamp", "SKU"])
        self.sku_table.horizontalHeader().setStretchLastSection(True)
        self.sku_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.sku_table.setEditTriggers(QTableWidget.NoEditTriggers)

        # Layouts
        controls_layout = QVBoxLayout()
        controls_layout.addWidget(QLabel("Select Serial Port:"))
        controls_layout.addWidget(self.serial_port_combo)
        controls_layout.addWidget(QLabel("Select Camera Index:"))
        controls_layout.addWidget(self.camera_combo)
        controls_layout.addWidget(QLabel("Order Number:"))
        controls_layout.addWidget(self.order_input)
        controls_layout.addWidget(self.start_button)
        controls_layout.addWidget(self.stop_button)
        controls_layout.addWidget(self.snapshot_button)
        controls_layout.addWidget(self.clear_button)
        controls_layout.addWidget(self.select_folder_button)
        controls_layout.addWidget(self.export_button)
        controls_layout.addStretch()
        controls_layout.addWidget(self.barcode_label)
        controls_layout.addWidget(self.count_label)
        controls_layout.addWidget(self.status_label)

        right_widget = QWidget()
        right_widget.setLayout(controls_layout)

        main_layout = QHBoxLayout()
        main_layout.addWidget(self.image_label, 2)
        main_layout.addWidget(right_widget, 1)

        # Combine main layout + table
        main_vlayout = QVBoxLayout()
        main_vlayout.addLayout(main_layout)
        main_vlayout.addWidget(self.sku_table)

        main_widget = QWidget()
        main_widget.setLayout(main_vlayout)

        # --- Settings Tab ---
        self.beep_checkbox = QCheckBox("Enable Beep Sound")
        self.beep_checkbox.setChecked(self.settings.get("beep_enabled", True))

        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["Dark", "Light"])
        self.theme_combo.setCurrentText(self.settings.get("theme", "Dark"))

        settings_layout = QVBoxLayout()
        settings_layout.addWidget(self.beep_checkbox)
        settings_layout.addWidget(QLabel("Theme:"))
        settings_layout.addWidget(self.theme_combo)

        self.save_settings_button = QPushButton("Save Settings")
        settings_layout.addWidget(self.save_settings_button)

        settings_widget = QWidget()
        settings_widget.setLayout(settings_layout)

        self.tabs.addTab(main_widget, "Scanner")
        self.tabs.addTab(settings_widget, "Settings")
        self.setCentralWidget(self.tabs)

        # --- Menu Bar ---
        menubar = self.menuBar()

        # File Menu
        file_menu = menubar.addMenu("File")
        open_action = file_menu.addAction("Open Folder")
        exit_action = file_menu.addAction("Exit")
        open_action.setShortcut("Ctrl+O")
        exit_action.setShortcut("Ctrl+Q")
        open_action.triggered.connect(self.select_folder)
        exit_action.triggered.connect(self.close)

        # Help Menu
        help_menu = menubar.addMenu("Help")
        about_action = help_menu.addAction("About")
        about_action.setShortcut("F1")
        about_action.triggered.connect(self.show_about)

        # --- Button Connections ---
        self.start_button.clicked.connect(self.start_camera)
        self.stop_button.clicked.connect(self.stop_camera)
        self.snapshot_button.clicked.connect(self.capture_snapshot)
        self.clear_button.clicked.connect(self.clear_session)
        self.select_folder_button.clicked.connect(self.select_folder)
        self.export_button.clicked.connect(self.export_to_excel)
        self.save_settings_button.clicked.connect(self.save_settings)

    # ------------------------ Menu & About ------------------------
    def show_about(self):
        QMessageBox.information(
            self,
            "About Barcode Inspector",
            "Barcode Inspector v1.0\n\n"
            "Scan, save, and export barcodes.\n"
            "Developed with PyQt5 and OpenCV.\n"
            "Developed by Hassan Irfan."
        )

    # ------------------------ Camera & Serial ------------------------
    def populate_serial_ports(self):
        ports = serial.tools.list_ports.comports()
        self.serial_port_combo.clear()
        self.serial_port_combo.addItem("None")
        for port in ports:
            self.serial_port_combo.addItem(port.device)
        self.serial_port_combo.setCurrentText("None")

    def _get_preferred_backend(self):
        """
        Return the preferable OpenCV backend flag for the current OS.
        On Linux prefer V4L2, on Windows prefer DSHOW, fallback to default (0).
        """
        if platform.system() == "Linux":
            return getattr(cv2, "CAP_V4L2", 0)
        if platform.system() == "Windows":
            return getattr(cv2, "CAP_DSHOW", 0)
        # macOS or others: use default
        return 0

    def detect_cameras(self, max_test=8):
        """
        Scan camera indexes and return a list of available indexes (strings).
        Uses the preferred backend for the platform, then fallback to default.
        """
        found = []
        backend = self._get_preferred_backend()
        for index in range(max_test):
            try:
                if backend:
                    cap = cv2.VideoCapture(index, backend)
                else:
                    cap = cv2.VideoCapture(index)
                # small delay not necessary usually but some devices need time
                if cap is not None and cap.isOpened():
                    found.append(str(index))
                    cap.release()
            except Exception:
                # ignore camera probe errors
                pass
        return found

    def populate_camera_indices(self, max_test=8):
        """
        Populate the camera_combo with valid camera indices detected.
        If none found, still add '0' so user can try manually.
        """
        self.camera_combo.clear()
        cameras = self.detect_cameras(max_test=max_test)
        if cameras:
            self.camera_combo.addItems(cameras)
            # select first found by default
            self.camera_combo.setCurrentIndex(0)
            self.status_label.setText(f"Status: {len(cameras)} camera(s) found")
        else:
            # no cameras discovered; add 0 as fallback option
            self.camera_combo.addItem("0")
            self.camera_combo.setCurrentText("0")
            self.status_label.setText("Status: No cameras auto-detected (try index 0,1,2...)")

    def _open_camera_with_fallback(self, idx):
        """
        Try opening camera idx using multiple strategies:
          1) preferred backend (V4L2 / DSHOW)
          2) default backend
        Returns cv2.VideoCapture if opened, else None.
        """
        # try preferred
        backend = self._get_preferred_backend()
        attempts = []

        def try_open(index, backend_flag=None):
            try:
                if backend_flag:
                    cap = cv2.VideoCapture(index, backend_flag)
                else:
                    cap = cv2.VideoCapture(index)
                if cap is not None and cap.isOpened():
                    return cap
                # ensure release if opened false
                if cap is not None:
                    cap.release()
            except Exception:
                pass
            return None

        # 1) try preferred backend
        if backend:
            attempts.append(("preferred", backend))
            cap = try_open(idx, backend)
            if cap:
                return cap

        # 2) try default
        attempts.append(("default", None))
        cap = try_open(idx, None)
        if cap:
            return cap

        # 3) try DSHOW on Windows if preferred wasn't that
        if platform.system() == "Windows":
            fallback_backend = getattr(cv2, "CAP_DSHOW", None)
            if fallback_backend and fallback_backend != backend:
                attempts.append(("dshow", fallback_backend))
                cap = try_open(idx, fallback_backend)
                if cap:
                    return cap

        # 4) try V4L2 on Linux if preferred wasn't that
        if platform.system() == "Linux":
            fallback_backend = getattr(cv2, "CAP_V4L2", None)
            if fallback_backend and fallback_backend != backend:
                attempts.append(("v4l2", fallback_backend))
                cap = try_open(idx, fallback_backend)
                if cap:
                    return cap

        # nothing worked
        return None

    def start_camera(self):
        serial_selection = self.serial_port_combo.currentText()
        if serial_selection is not None and serial_selection != "None":
            try:
                self.serial_port = serial.Serial(serial_selection, 9600, timeout=1)
                self.status_label.setText(f"Connected to {serial_selection}")
            except Exception as e:
                self.show_error(f"Serial error: {e}")
                return
        else:
            self.serial_port = None
            # do not overwrite status if camera status exists

        # Try to read index from combo; if invalid, fallback to scanning
        selected_text = ""
        try:
            selected_text = self.camera_combo.currentText() if self.camera_combo.currentText() is not None else ""
        except Exception:
            selected_text = ""

        # If the combo is empty or contains a non-integer, attempt to auto-find
        use_index = None
        if selected_text and selected_text.isdigit():
            use_index = int(selected_text)
        else:
            # attempt to auto-find the first working camera
            detected = self.detect_cameras(max_test=8)
            if detected:
                use_index = int(detected[0])
            else:
                # fallback to brute-force scanning 0..8
                use_index = None

        cap = None
        if use_index is not None:
            cap = self._open_camera_with_fallback(use_index)
            if not cap:
                # maybe the camera is at different index; try scanning 0..8
                cap = self._scan_and_open_any_camera(max_scan=8)
        else:
            # try scanning 0..8 to find any working camera
            cap = self._scan_and_open_any_camera(max_scan=8)

        if not cap:
            self.show_error("Failed to open any camera. Make sure camera isn't used by another app and permissions are allowed.")
            self.status_label.setText("Status: Camera open failed")
            return

        # success
        self.capture = cap
        # set resolution (some cameras ignore this)
        try:
            self.capture.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
            self.capture.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)
        except Exception:
            pass

        # store camera_index for reference
        try:
            self.camera_index = int(use_index) if use_index is not None else None
        except Exception:
            self.camera_index = None

        # Load previous SKUs if folder exists
        order_number = self.order_input.text().strip() or "NoOrder"
        if self.save_location:
            self.load_existing_barcodes(order_number)

        self.timer.start(30)
        self.status_label.setText(f"Status: Camera started (index {self.camera_index if self.camera_index is not None else 'unknown'})")

    def _scan_and_open_any_camera(self, max_scan=8):
        """
        Scan indexes 0..max_scan-1 and try to open them until one succeeds.
        Returns the opened cv2.VideoCapture or None.
        """
        for idx in range(max_scan):
            cap = self._open_camera_with_fallback(idx)
            if cap:
                # update combo to show the index we found
                try:
                    idx_str = str(idx)
                    # if this index isn't in combo, add it
                    if self.camera_combo.findText(idx_str) == -1:
                        self.camera_combo.addItem(idx_str)
                    self.camera_combo.setCurrentText(idx_str)
                except Exception:
                    pass
                return cap
        return None

    def stop_camera(self):
        self.timer.stop()
        if self.capture:
            try:
                self.capture.release()
            except Exception:
                pass
            self.capture = None
        if self.serial_port:
            try:
                self.serial_port.close()
            except Exception:
                pass
            self.serial_port = None
        self.image_label.setText("Camera Feed")
        self.status_label.setText("Status: Disconnected")

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Save Folder")
        if folder:
            self.save_location = folder
            QMessageBox.information(self, "Folder Selected", f"Save folder: {folder}")

    # ------------------------ Frame Processing ------------------------
    def update_frame(self):
        if not self.capture:
            return
        try:
            ret, frame = self.capture.read()
        except Exception:
            ret = False
            frame = None

        if not ret or frame is None:
            # camera stopped or returned empty frame
            # stop and show error once
            self.stop_camera()
            self.show_error("Camera disconnected or returned no frames.")
            return

        # decode barcodes
        try:
            barcodes = pyzbar.decode(frame, symbols=[ZBarSymbol.EAN13, ZBarSymbol.CODE128, ZBarSymbol.QRCODE, ZBarSymbol.DATAMATRIX])
        except Exception:
            barcodes = []

        current_barcode = None

        for barcode in barcodes:
            try:
                barcode_data = barcode.data.decode("utf-8").strip()
            except Exception:
                continue
            current_barcode = barcode_data

            # draw polygon
            try:
                pts = np.array([barcode.polygon], np.int32)
                color = (0, 0, 255) if barcode_data == self.last_barcode else (0, 255, 0)
                cv2.polylines(frame, [pts], True, color, 3)
                x, y, w, h = cv2.boundingRect(np.array(barcode.polygon))
                cv2.putText(frame, barcode_data, (x, y - 10),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.8, color, 2)
            except Exception:
                pass

            sku_id = self.extract_sku(barcode_data)
            if sku_id not in self.barcode_set:
                self.barcode_set.add(sku_id)
                self.capture_image(frame, sku_id)
                self.last_barcode = barcode_data
                self.add_to_table(time.strftime("%Y-%m-%d %H:%M:%S"), sku_id)
                if self.serial_port:
                    try:
                        self.serial_port.write((barcode_data + '\n').encode())
                    except Exception:
                        # if serial write fails, play error beep (if enabled)
                        if self.settings.get("beep_enabled", True):
                            self.play_sound(success=False)

        # Update labels
        if current_barcode:
            self.barcode_label.setText(f"Current Barcode: {current_barcode}")
            self.count_label.setText(f"Unique SKUs Scanned: {len(self.barcode_set)}")
            if len(self.barcode_set) > 50:
                self.count_label.setStyleSheet("color: red; font-weight: bold;")
            else:
                self.count_label.setStyleSheet("color: lime; font-weight: bold;")
        else:
            self.barcode_label.setText("Current Barcode: None")

        # Convert frame to Qt format
        try:
            rgb_image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            h, w, ch = rgb_image.shape
            bytes_per_line = ch * w
            qt_image = QImage(rgb_image.data, w, h, bytes_per_line, QImage.Format_RGB888)
            scaled_image = QPixmap.fromImage(qt_image).scaled(
                self.image_label.width(), self.image_label.height(),
                Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.image_label.setPixmap(scaled_image)
        except Exception:
            # If conversion fails, just show placeholder text
            self.image_label.setText("Camera Feed - frame conversion failed")

    # ------------------------ Core Logic ------------------------
    def extract_sku(self, barcode_value):
        try:
            return barcode_value.split()[0].split('-')[0]
        except Exception:
            return barcode_value

    def capture_image(self, frame, sku_value):
        if not self.save_location:
            # if user didn't choose a folder, fallback to current working dir
            self.save_location = os.path.abspath(".")
            # notify user once
            QMessageBox.information(self, "Save Folder", f"No folder selected — using {self.save_location}")

        order_number = self.order_input.text().strip() or "NoOrder"
        order_folder = os.path.join(self.save_location, order_number)
        os.makedirs(order_folder, exist_ok=True)

        timestamp = time.strftime("%Y%m%d-%H%M%S")
        filename = os.path.join(order_folder, f"{sku_value}_{timestamp}.jpg")
        try:
            cv2.imwrite(filename, frame)
        except Exception as e:
            # save failed: show error but continue
            print("Failed to write image:", e)

        if self.settings.get("beep_enabled", True) and self.beep_checkbox.isChecked():
            self.play_sound(success=True)

        log_path = os.path.join(order_folder, "barcode_log.csv")
        file_exists = os.path.exists(log_path)
        try:
            with open(log_path, "a") as f:
                if not file_exists:
                    f.write("Timestamp,SKU\n")
                f.write(f"{timestamp},{sku_value}\n")
        except Exception as e:
            print("Failed to write log:", e)

    def capture_snapshot(self):
        if self.capture:
            try:
                ret, frame = self.capture.read()
            except Exception:
                ret = False
                frame = None
            if ret and frame is not None:
                self.capture_image(frame, "manual")
            else:
                self.show_error("Failed to capture snapshot (camera not returning frames).")

    def export_to_excel(self):
        if not self.save_location:
            self.show_error("No save folder selected")
            return

        order_number = self.order_input.text().strip() or "NoOrder"
        log_path = os.path.join(self.save_location, order_number, "barcode_log.csv")
        if os.path.exists(log_path) and os.path.getsize(log_path) > 0:
            df = pd.read_csv(log_path)
            if not df.empty:
                excel_path = os.path.join(self.save_location, order_number, "barcode_log.xlsx")
                try:
                    df.to_excel(excel_path, index=False)
                    QMessageBox.information(self, "Export", f"Log exported to {excel_path}")
                    return
                except Exception as e:
                    self.show_error(f"Failed to export to Excel: {e}")
                    return
        self.show_error("No valid barcode log to export")

    def clear_session(self):
        self.barcode_set.clear()
        self.last_barcode = None
        self.barcode_label.setText("Current Barcode: None")
        self.count_label.setText("Unique SKUs Scanned: 0")
        self.count_label.setStyleSheet("color: lime; font-weight: bold;")
        self.sku_table.setRowCount(0)
        QMessageBox.information(self, "Session", "Session cleared successfully!")

    # ------------------------ Table Helpers ------------------------
    def add_to_table(self, timestamp, sku):
        row = self.sku_table.rowCount()
        self.sku_table.insertRow(row)
        self.sku_table.setItem(row, 0, QTableWidgetItem(timestamp))
        self.sku_table.setItem(row, 1, QTableWidgetItem(sku))

    # ------------------------ Helpers ------------------------
    def play_sound(self, success=True):
        def beep():
            try:
                if platform.system() == "Windows":
                    import winsound
                    freq, dur = (1000, 150) if success else (400, 300)
                    winsound.Beep(freq, dur)
                else:
                    freq = 1000 if success else 400
                    duration = 0.15 if success else 0.3
                    # use 'play' if installed; otherwise try 'aplay' or fallback silent
                    if shutil_which("play"):
                        os.system(f"play -nq -t alsa synth {duration} sine {freq} 2>/dev/null")
                    elif shutil_which("aplay"):
                        # simple beep using /dev/zero won't produce tone; skipping
                        pass
            except Exception:
                pass

        threading.Thread(target=beep, daemon=True).start()

    def show_error(self, message):
        QMessageBox.critical(self, "Error", message)

    # ------------------------ Settings ------------------------
    def load_settings(self):
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, 'r') as f:
                    self.settings = json.load(f)
            except Exception:
                self.settings = {}
        else:
            self.settings = {}

    def save_settings(self):
        self.settings["beep_enabled"] = self.beep_checkbox.isChecked()
        self.settings["theme"] = self.theme_combo.currentText()
        try:
            with open(self.settings_file, 'w') as f:
                json.dump(self.settings, f)
        except Exception:
            pass
        self.set_stylesheet()
        QMessageBox.information(self, "Settings", "Settings saved successfully!")

    def set_stylesheet(self):
        theme = self.settings.get("theme", "Dark")
        if theme == "Dark":
            self.setStyleSheet("""
                QWidget { background-color: #121212; color: white; }
                QPushButton { font-size: 14px; padding: 6px; border-radius: 5px; }
                QLabel, QComboBox, QLineEdit { color: white; font-size: 16px; }
                QTableWidget { background-color: #1e1e1e; color: white; gridline-color: gray; }
            """)
        else:
            self.setStyleSheet("""
                QWidget { background-color: white; color: black; }
                QPushButton { font-size: 14px; padding: 6px; border-radius: 5px; }
                QLabel, QComboBox, QLineEdit { color: black; font-size: 16px; }
                QTableWidget { background-color: #f0f0f0; color: black; gridline-color: gray; }
            """)

    def load_existing_barcodes(self, order_number):
        self.barcode_set.clear()
        log_path = os.path.join(self.save_location, order_number, "barcode_log.csv")
        if os.path.exists(log_path):
            try:
                df = pd.read_csv(log_path)
                self.barcode_set.update(df['SKU'].astype(str).tolist())
                # Populate table
                self.sku_table.setRowCount(0)
                for i, row in df.iterrows():
                    self.add_to_table(row['Timestamp'], row['SKU'])
            except Exception:
                pass

# small helper used in play_sound
def shutil_which(cmd):
    from shutil import which
    return which(cmd)

# ------------------------ Main Entry ------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = BarcodeApp()
    window.show()
    sys.exit(app.exec_())

