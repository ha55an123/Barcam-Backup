#!/usr/bin/env python3
"""
Barcode Inspector - webcam fixes and enhanced camera handling
"""

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
import shutil
from PyQt5.QtWidgets import (
    QApplication, QLabel, QPushButton, QVBoxLayout, QHBoxLayout, QWidget, QFileDialog,
    QComboBox, QMessageBox, QMainWindow, QLineEdit, QTabWidget, QCheckBox, QTableWidget,
    QTableWidgetItem, QHeaderView
)
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtGui import QImage, QPixmap

# ---------- Barcode Decoders ----------
try:
    from pyzbar import pyzbar
    from pyzbar.pyzbar import ZBarSymbol
    PYZBAR_AVAILABLE = True
except Exception as e:
    PYZBAR_AVAILABLE = False
    print("pyzbar import failed:", repr(e))

try:
    from pylibdmtx.pylibdmtx import decode as dmtx_decode
    DMTX_AVAILABLE = True
except Exception as e:
    DMTX_AVAILABLE = False
    print("pylibdmtx not available:", repr(e))


class BarcodeApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Barcode Inspector")
        self.setGeometry(100, 100, 1300, 800)

        # Load settings
        self.settings_file = "settings.json"
        self.settings = {}
        self.load_settings()

        # Camera and serial
        self.capture = None
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_frame)
        self.camera_index = None
        self.last_barcode = None
        self.barcode_set = set()
        self.save_location = ""
        self.serial_port = None

        self.init_ui()
        self.populate_serial_ports()
        self.populate_camera_indices()
        self.set_stylesheet()

        decoders = []
        if PYZBAR_AVAILABLE:
            decoders.append("pyzbar")
        if DMTX_AVAILABLE:
            decoders.append("pylibdmtx")
        self.status_label.setText(
            "Status: Ready — decoders: " + ", ".join(decoders) if decoders else
            "Status: No barcode decoders available"
        )

    # ------------------------ UI ------------------------
    def init_ui(self):
        self.tabs = QTabWidget()
        # Scanner Tab
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

        # Status label
        self.status_label = QLabel("Status: Initializing...")

        # Table
        self.sku_table = QTableWidget()
        self.sku_table.setColumnCount(2)
        self.sku_table.setHorizontalHeaderLabels(["Timestamp", "SKU"])
        self.sku_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.sku_table.setEditTriggers(QTableWidget.NoEditTriggers)

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
        controls_layout.addWidget(self.barcode_label)
        controls_layout.addWidget(self.count_label)
        controls_layout.addWidget(self.status_label)

        right_widget = QWidget()
        right_widget.setLayout(controls_layout)

        main_layout = QHBoxLayout()
        main_layout.addWidget(self.image_label, 2)
        main_layout.addWidget(right_widget, 1)

        main_vlayout = QVBoxLayout()
        main_vlayout.addLayout(main_layout)
        main_vlayout.addWidget(self.sku_table)

        main_widget = QWidget()
        main_widget.setLayout(main_vlayout)

        # Settings Tab
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

        # Connections
        self.start_button.clicked.connect(self.start_camera)
        self.stop_button.clicked.connect(self.stop_camera)
        self.snapshot_button.clicked.connect(self.capture_snapshot)
        self.clear_button.clicked.connect(self.clear_session)
        self.select_folder_button.clicked.connect(self.select_folder)
        self.export_button.clicked.connect(self.export_to_excel)
        self.save_settings_button.clicked.connect(self.save_settings)

    # ------------------------ Camera ------------------------
    def _get_backend(self):
        if platform.system() == "Linux":
            return cv2.CAP_V4L2
        if platform.system() == "Windows":
            return cv2.CAP_DSHOW
        return cv2.CAP_ANY

    def detect_cameras(self, max_test=5):
        found = []
        backend = self._get_backend()
        for i in range(max_test):
            cap = cv2.VideoCapture(i, backend)
            if cap.isOpened():
                found.append(str(i))
                cap.release()
        return found

    def populate_camera_indices(self):
        self.camera_combo.clear()
        cams = self.detect_cameras()
        if cams:
            self.camera_combo.addItems(cams)
            self.camera_combo.setCurrentIndex(0)
            self.status_label.setText(f"Status: {len(cams)} camera(s) found")
        else:
            self.camera_combo.addItem("0")
            self.camera_combo.setCurrentText("0")
            self.status_label.setText("Status: No cameras detected")

    def _open_camera(self, idx):
        backend = self._get_backend()
        cap = cv2.VideoCapture(idx, backend)
        if cap and cap.isOpened():
            return cap
        # fallback
        cap.release()
        cap = cv2.VideoCapture(idx)
        if cap and cap.isOpened():
            return cap
        return None

    def start_camera(self):
        selected = self.camera_combo.currentText()
        index = int(selected) if selected.isdigit() else 0
        cap = self._open_camera(index)
        if not cap:
            self.status_label.setText("Status: Failed to open camera")
            self.show_error("Could not open camera. Try different index or check permissions.")
            return
        self.capture = cap
        self.capture.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
        self.capture.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)
        self.timer.start(30)
        self.status_label.setText(f"Status: Camera started (index {index})")

    def stop_camera(self):
        self.timer.stop()
        if self.capture:
            self.capture.release()
            self.capture = None
        self.image_label.setText("Camera Feed")
        self.status_label.setText("Status: Camera stopped")

    # ------------------------ Frame Update ------------------------
    def update_frame(self):
        if not self.capture:
            return
        ret, frame = self.capture.read()
        if not ret or frame is None:
            return
        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        h, w, ch = rgb.shape
        bytes_per_line = ch * w
        qt_image = QImage(rgb.data, w, h, bytes_per_line, QImage.Format_RGB888)
        self.image_label.setPixmap(QPixmap.fromImage(qt_image).scaled(
            self.image_label.width(), self.image_label.height(),
            Qt.KeepAspectRatio, Qt.SmoothTransformation
        ))

    # ------------------------ Helpers ------------------------
    def load_settings(self):
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, "r") as f:
                    self.settings = json.load(f)
            except:
                self.settings = {}
        else:
            self.settings = {}

    def save_settings(self):
        self.settings["beep_enabled"] = self.beep_checkbox.isChecked()
        self.settings["theme"] = self.theme_combo.currentText()
        with open(self.settings_file, "w") as f:
            json.dump(self.settings, f)
        self.set_stylesheet()

    def set_stylesheet(self):
        theme = self.settings.get("theme", "Dark")
        if theme == "Dark":
            self.setStyleSheet("QWidget { background-color: #121212; color: white; }")
        else:
            self.setStyleSheet("QWidget { background-color: white; color: black; }")

    def populate_serial_ports(self):
        ports = [p.device for p in serial.tools.list_ports.comports()]
        self.serial_port_combo.clear()
        self.serial_port_combo.addItem("None")
        self.serial_port_combo.addItems(ports)

    def show_error(self, msg):
        QMessageBox.critical(self, "Error", msg)

    def capture_snapshot(self):
        if self.capture:
            ret, frame = self.capture.read()
            if ret:
                filename = os.path.join(os.getcwd(), f"snapshot_{int(time.time())}.jpg")
                cv2.imwrite(filename, frame)
                QMessageBox.information(self, "Snapshot", f"Saved {filename}")
            else:
                self.show_error("Failed to capture snapshot")

# ------------------------ Main ------------------------
def main():
    app = QApplication(sys.argv)
    window = BarcodeApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()

