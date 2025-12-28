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
import subprocess

from PyQt5.QtWidgets import (
    QApplication, QLabel, QPushButton, QVBoxLayout, QWidget, QFileDialog,
    QComboBox, QMessageBox, QMainWindow, QLineEdit, QTabWidget, QCheckBox,
    QHBoxLayout
)
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtGui import QImage, QPixmap
from pyzbar.pyzbar import decode
class BarcodeApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Barcode Inspector")
        self.setGeometry(100, 100, 1000, 800)

        self.settings_file = "settings.json"
        self.load_settings()

        self.init_ui()
        self.populate_serial_ports()
        self.populate_camera_indices()

        self.capture = None
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_frame)
        self.camera_index = 0

        self.last_barcode = None
        self.barcode_set = set()
        self.save_location = ""
        self.serial_port = None

        self.set_stylesheet()

    def init_ui(self):
        self.tabs = QTabWidget()

        # Main Tab
        self.image_label = QLabel("Camera Feed")
        self.image_label.setFixedSize(500, 300)
        self.barcode_label = QLabel("Current Barcode: None")
        self.count_label = QLabel("Barcodes Scanned: 0")

        self.serial_port_combo = QComboBox()
        self.camera_combo = QComboBox()
        self.start_button = QPushButton("Start Camera")
        self.stop_button = QPushButton("Stop Camera")
        self.snapshot_button = QPushButton("Capture Snapshot")
        self.select_folder_button = QPushButton("Select Save Folder")
        self.export_button = QPushButton("Export to Excel")
        self.order_input = QLineEdit()
        self.order_input.setPlaceholderText("Enter Order Number")

        self.status_label = QLabel("Status: Not Connected")

        layout = QVBoxLayout()
        layout.addWidget(self.image_label)
        layout.addWidget(self.barcode_label)
        layout.addWidget(self.count_label)
        layout.addWidget(QLabel("Select Serial Port:"))
        layout.addWidget(self.serial_port_combo)
        layout.addWidget(QLabel("Select Camera Index:"))
        layout.addWidget(self.camera_combo)
        layout.addWidget(QLabel("Order Number:"))
        layout.addWidget(self.order_input)
        layout.addWidget(self.start_button)
        layout.addWidget(self.stop_button)
        layout.addWidget(self.snapshot_button)
        layout.addWidget(self.select_folder_button)
        layout.addWidget(self.export_button)
        layout.addWidget(self.status_label)

        main_widget = QWidget()
        main_widget.setLayout(layout)

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

        self.start_button.clicked.connect(self.start_camera)
        self.stop_button.clicked.connect(self.stop_camera)
        self.snapshot_button.clicked.connect(self.capture_snapshot)
        self.select_folder_button.clicked.connect(self.select_folder)
        self.export_button.clicked.connect(self.export_to_excel)
        self.save_settings_button.clicked.connect(self.save_settings)

    def populate_serial_ports(self):
        ports = serial.tools.list_ports.comports()
        self.serial_port_combo.clear()
        for port in ports:
            self.serial_port_combo.addItem(port.device)

    def get_camera_name(self, index):
        # Linux only: try getting camera name via v4l2-ctl
        dev = f"/dev/video{index}"
        if os.path.exists(dev):
            try:
                result = subprocess.check_output(["v4l2-ctl", "-d", dev, "--info"]).decode()
                for line in result.splitlines():
                    if "Driver name" in line:
                        return f"{index} - {line.strip()}"
            except Exception:
                pass
        return f"{index} - Camera"

    def populate_camera_indices(self, max_test=10):
        self.camera_combo.clear()
        for index in range(max_test):
            cap = cv2.VideoCapture(index)
            if cap.isOpened():
                name = self.get_camera_name(index)
                self.camera_combo.addItem(name)
                cap.release()

    def start_camera(self):
        try:
            self.serial_port = serial.Serial(self.serial_port_combo.currentText(), 9600, timeout=1)
            self.status_label.setText("Connected to serial.")
        except Exception as e:
            self.show_error(f"Serial error: {e}")
            return

        try:
            selected_text = self.camera_combo.currentText()
            self.camera_index = int(selected_text.split(" - ")[0])
        except ValueError:
            self.show_error("Invalid camera index.")
            return

        self.capture = cv2.VideoCapture(self.camera_index)
        self.capture.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
        self.capture.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)

        if not self.capture.isOpened():
            self.show_error("Failed to open camera.")
            return

        self.timer.start(30)

    def stop_camera(self):
        self.timer.stop()
        if self.capture:
            self.capture.release()
        if self.serial_port:
            self.serial_port.close()
        self.image_label.setText("Camera Feed")
        self.barcode_label.setText("Current Barcode: None")
        self.count_label.setText("Barcodes Scanned: 0")
        self.barcode_set.clear()
        self.last_barcode = None
        self.status_label.setText("Status: Disconnected")

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Save Folder")
        if folder:
            self.save_location = folder

    def update_frame(self):
        ret, frame = self.capture.read()
        if not ret:
            return

        barcodes = decode(frame)
        current_barcode = None

        for barcode in barcodes:
            barcode_data = barcode.data.decode("utf-8")
            current_barcode = barcode_data

            points = barcode.polygon
            if len(points) == 4:
                pts = [tuple(point) for point in points]
                cv2.polylines(frame, [np.array(pts, dtype=np.int32)], True, (0, 255, 0), 3)
            else:
                rect = barcode.rect
                cv2.rectangle(frame, (rect[0], rect[1]),
                              (rect[0] + rect[2], rect[1] + rect[3]), (0, 255, 0), 3)

            cv2.putText(frame, current_barcode, (barcode.rect[0], barcode.rect[1] - 10),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.8, (255, 0, 0), 2)

            if current_barcode != self.last_barcode:
                self.last_barcode = current_barcode
                self.capture_image(frame, current_barcode)

        if current_barcode:
            self.barcode_label.setText(f"Current Barcode: {current_barcode}")

        rgb_image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        h, w, ch = rgb_image.shape
        bytes_per_line = ch * w
        qt_image = QImage(rgb_image.data, w, h, bytes_per_line, QImage.Format_RGB888)
        scaled_image = QPixmap.fromImage(qt_image).scaled(
            self.image_label.width(), self.image_label.height(),
            Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.image_label.setPixmap(scaled_image)

    def capture_image(self, frame, barcode_value):
        if not self.save_location:
            self.show_error("No save location selected!")
            return

        order_number = self.order_input.text().strip() or "NoOrder"
        order_folder = os.path.join(self.save_location, order_number)
        os.makedirs(order_folder, exist_ok=True)

        timestamp = time.strftime("%Y%m%d-%H%M%S")
        filename = os.path.join(order_folder, f"barcode_{barcode_value}_{timestamp}.jpg")
        cv2.imwrite(filename, frame)

        if self.beep_checkbox.isChecked():
            self.play_sound()

        log_path = os.path.join(order_folder, "barcode_log.csv")
        with open(log_path, "a") as f:
            f.write(f"{timestamp},{barcode_value}\n")

        self.barcode_set.add(barcode_value)
        self.count_label.setText(f"Barcodes Scanned: {len(self.barcode_set)}")

    def capture_snapshot(self):
        if self.capture:
            ret, frame = self.capture.read()
            if ret:
                self.capture_image(frame, "manual")

    def export_to_excel(self):
        if not self.save_location:
            self.show_error("No save folder selected")
            return

        order_number = self.order_input.text().strip() or "NoOrder"
        log_path = os.path.join(self.save_location, order_number, "barcode_log.csv")
        if os.path.exists(log_path):
            df = pd.read_csv(log_path)
            excel_path = os.path.join(self.save_location, order_number, "barcode_log.xlsx")
            df.to_excel(excel_path, index=False)
            QMessageBox.information(self, "Export", f"Log exported to {excel_path}")
        else:
            self.show_error("No barcode log to export")

    def play_sound(self):
        def beep():
            if platform.system() == "Windows":
                import winsound
                winsound.Beep(1000, 150)
            else:
                os.system("play -nq -t alsa synth 0.15 sine 1000")

        threading.Thread(target=beep).start()

    def show_error(self, message):
        QMessageBox.critical(self, "Error", message)

    def load_settings(self):
        if os.path.exists(self.settings_file):
            with open(self.settings_file, 'r') as f:
                self.settings = json.load(f)
        else:
            self.settings = {}

    def save_settings(self):
        self.settings["beep_enabled"] = self.beep_checkbox.isChecked()
        self.settings["theme"] = self.theme_combo.currentText()
        with open(self.settings_file, 'w') as f:
            json.dump(self.settings, f)
        self.set_stylesheet()
        QMessageBox.information(self, "Settings", "Settings saved successfully!")

    def set_stylesheet(self):
        theme = self.settings.get("theme", "Dark")
        if theme == "Dark":
            self.setStyleSheet("""
                QWidget {
                    background-color: black;
                    color: white;
                }
                QPushButton {
                    background-color: grey;
                    color: white;
                    font-size: 14px;
                    padding: 5px;
                    border-radius: 5px;
                }
                QLabel, QComboBox, QLineEdit {
                    color: white;
                    font-size: 16px;
                }
            """)
        else:
            self.setStyleSheet("")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = BarcodeApp()
    window.show()
    sys.exit(app.exec_())
