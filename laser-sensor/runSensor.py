import sys
import time
import threading
from queue import Queue, Empty

from PyQt5 import QtWidgets, QtCore
import pyqtgraph as pg
import serial
import serial.tools.list_ports

import os
import csv
from datetime import datetime
import pandas as pd
from PyQt5.QtWidgets import QFileDialog, QMessageBox

class SerialThread(threading.Thread):
    """
    Thread that reads COM3 at 9600 baud and enqueues ES and EN values.
    """
    def __init__(self, port='COM3', baud=9600):
        super().__init__(daemon=True)
        self.port = port
        self.baud = baud
        self.is_collecting = False
        self._stop = False
        self.queue = Queue()
        self.markers = ['EN', 'ES', 'NE', 'NW', 'WN', 'SE', 'SW', 'WS']
    def run(self):
        try:
            ser = serial.Serial(self.port, self.baud, timeout=1)
            print(f"[SerialThread] Opened {self.port} at {self.baud} baud.")
        except Exception as e:
            print(f"[SerialThread] ERROR opening {self.port}: {e}")
            return

        while not self._stop:
            if self.is_collecting and ser.in_waiting:
                line = ser.readline().decode('utf-8', errors='ignore').strip()
                if line:
                    # parse out ALL SENSOR values
                    parts = line.split('|')
                    ts = time.time()
                    for chunk in parts:
                        for marker in self.markers:
                            if chunk.startswith(marker):
                                try:
                                    val = float(chunk[len(marker):])
                                    self.queue.put((marker, ts, val))
                                except ValueError:
                                    pass
            else:
                time.sleep(0.05)

        ser.close()
        print("[SerialThread] Serial connection closed.")

    def stop(self):
        self._stop = True


class MainWindow(QtWidgets.QMainWindow):
    """
    GUI with two real-time plots for ES and EN displacement.
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Laser Sensors Plots")
        self.resize(1000, 900)

        cw = QtWidgets.QWidget()
        self.setCentralWidget(cw)
        layout = QtWidgets.QVBoxLayout(cw)
        
        self.markers = ['EN', 'ES', 'NE', 'NW', 'WN', 'SE', 'SW', 'WS']
        self.part_data = {marker: [] for marker in self.markers}
        self.plots = {}
        self.curves = {}

        grid = QtWidgets.QGridLayout()
        for idx, marker in enumerate(self.markers):
            plot = pg.PlotWidget(title=f"{marker}")
            plot.setLabel('bottom', 'Time (s)')
            plot.setLabel('left',   'Displacement (in)')
            curve = plot.plot([], [], pen=pg.mkPen(width=2))
            self.plots[marker] = plot
            self.curves[marker] = curve
            grid.addWidget(plot, idx // 2, idx % 2)
        layout.addLayout(grid)

        # # ES plot
        # self.plot_es = pg.PlotWidget(title="ES Displacement")
        # self.plot_es.setLabel('bottom', 'Time Elapsed (s)')
        # self.plot_es.setLabel('left',   'ES Value (in)')
        # layout.addWidget(self.plot_es)
        # self.es_curve = self.plot_es.plot([], [], pen=pg.mkPen('y', width=2))
        # self.es_data = []

        # # EN plot
        # self.plot_en = pg.PlotWidget(title="EN Displacement")
        # self.plot_en.setLabel('bottom', 'Time Elapsed (s)')
        # self.plot_en.setLabel('left',   'EN Value (in)')
        # layout.addWidget(self.plot_en)
        # self.en_curve = self.plot_en.plot([], [], pen=pg.mkPen('r', width=2))
        # self.en_data = []

        # Buttons
        btn_layout = QtWidgets.QHBoxLayout()
        self.start_btn = QtWidgets.QPushButton("Start")
        self.stop_btn  = QtWidgets.QPushButton("Stop")
        self.save_btn  = QtWidgets.QPushButton("Save")  
        self.reset_btn = QtWidgets.QPushButton("Reset")  
        self.stop_btn.setEnabled(False)
        self.save_btn.setEnabled(False)  
        btn_layout.addWidget(self.start_btn)
        btn_layout.addWidget(self.stop_btn)
        btn_layout.addWidget(self.save_btn)  
        btn_layout.addWidget(self.reset_btn)  
        layout.addLayout(btn_layout)

        # Status
        self.status = QtWidgets.QLabel("Status: Idle")
        layout.addWidget(self.status)

        # Serial thread
        self.serial_thread = SerialThread(port='COM3', baud=9600) #this is HARDCODED
        self.serial_thread.start()

        # Timer to update both plots
        self.timer = QtCore.QTimer(self)
        self.timer.setInterval(100)
        self.timer.timeout.connect(self.update_plots)
        self.timer.start()

        # Signals
        self.start_btn.clicked.connect(self.start_recording)
        self.stop_btn.clicked.connect(self.stop_recording)
        self.save_btn.clicked.connect(self.save_to_excel)
        self.reset_btn.clicked.connect(self.reset_data)  

        self.first_start = True  

    def start_recording(self):
        
        if self.first_start:
            self.start_time = time.time()
            for marker in self.markers:
                self.part_data[marker].clear()
                self.curves[marker].setData([], [])
            self.status.setText("Status: Data initialized and recording")
            self.first_start = False
        else:
            self.status.setText("Status: Recording")
        self.serial_thread.is_collecting = True
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.save_btn.setEnabled(True)

    def stop_recording(self):
        self.serial_thread.is_collecting = False
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.status.setText("Status: Stopped")

    def reset_data(self):
        for marker in self.markers:
            self.part_data[marker].clear()
            self.curves[marker].setData([], [])
        self.status.setText("Status: Data Reset")
        self.first_start = True
        self.save_btn.setEnabled(False)  
    
    def save_to_excel(self):
        if not any(self.part_data[m] for m in self.markers):
            QMessageBox.warning(self, "Save Error", "No data to save.")
            return

        ts = datetime.now().strftime("%Y_%m_%d_%H%M%S")
        default_name = f"run_{ts}.xlsx"
        out_fn, _ = QFileDialog.getSaveFileName(self, "Save Excel File", default_name, "Excel Files (*.xlsx)")
        
        if not out_fn:
            return
    
        with pd.ExcelWriter(out_fn, engine='openpyxl') as writer:
            for m in self.markers:
                data = self.part_data[m]
                if data:
                    df = pd.DataFrame(data, columns=['Time (s)', 'Displacement (in)'])
                    df.to_excel(writer, sheet_name=m, index=False)
        QMessageBox.information(self, "Saved", f"All data exported to:\n{out_fn}")

    def update_plots(self):
        updated = set() #weird
        while True:
            try:
                sensor, ts, val = self.serial_thread.queue.get_nowait()
                t_rel = ts - self.start_time
                if sensor in self.part_data:
                    self.part_data[sensor].append((t_rel, val))
                    updated.add(sensor)
            except Empty:
                break

        for marker in updated:
            if self.part_data[marker]:
                x, y = zip(*self.part_data[marker])
                self.curves[marker].setData(x, y)

    def closeEvent(self, event):
        self.serial_thread.stop()
        super().closeEvent(event)


def main():
    """
    Run the application:
      pip install pyqt5 pyqtgraph pyserial
    """
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()


