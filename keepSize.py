import os
import sys
import time
import pathlib

os.environ['PATH'] = f"{os.environ['PATH']};{os.path.dirname(sys.executable)}\\Lib\\site-packages\\pywin32_system32"

import pywintypes
import win32gui
import win32con
import subprocess
from PyQt5 import QtCore
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTextEdit

home_dir = pathlib.Path.home()

# Create a subdirectory for your application's data files
data_dir = os.path.join(home_dir, '.keepsize')
os.makedirs(data_dir, exist_ok=True)

# Use this directory to save your data files
config_file = os.path.join(data_dir, 'config.txt')
if not os.path.exists(config_file):
    open(config_file, 'w')
window_sizes_file = os.path.join(data_dir, 'window_sizes.txt')
error_file = os.path.join(data_dir, 'errorlog.txt')

class KeepSize(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Keep Size')
        self.setGeometry(100, 100, 500, 400)

        self.save_button = QPushButton('Save', self)
        self.save_button.clicked.connect(self.save_config_file)

        self.grab_button = QPushButton('Grab Window Sizes', self)
        self.grab_button.clicked.connect(self.grab_window_sizes)

        self.set_button = QPushButton('Set Window Sizes', self)
        self.set_button.clicked.connect(self.set_window_sizes)

        self.text_box = QTextEdit(self)
        self.text_box.setPlainText(self.read_config_file())

        layout = QVBoxLayout()
        label = QLabel('You can edit the below configuration to add or remove windows that you want affected, \nsimply enter one word from the name of the window on a new line and click save.')
        label.setAlignment(QtCore.Qt.AlignCenter)
        layout.addWidget(label)
        layout.addWidget(self.text_box)
        layout.addWidget(self.save_button)  # moved this line before the QHBoxLayout
        button_layout = QHBoxLayout()
        layout.addLayout(button_layout)
        button_layout.addWidget(self.grab_button)
        button_layout.addWidget(self.set_button)
        self.setLayout(layout)

    def read_config_file(self):
        with open(config_file, "r") as file:
            return file.read()

    def save_config_file(self):
        with open(config_file, "w") as file:
            file.write(self.text_box.toPlainText())

    def get_window_sizes(self):
        sizes = []
        hwnds = []
        win32gui.EnumWindows(lambda hwnd, param: param.append(hwnds.append(hwnd)), [])
        for hwnd in hwnds:
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                for app in self.get_applications():
                    if app.lower() in title.lower():
                        rect = win32gui.GetWindowRect(hwnd)
                        width = rect[2] - rect[0]
                        height = rect[3] - rect[1]
                        left = rect[0]
                        top = rect[1]
                        sizes.append((title, app, width, height, left, top))
        return sizes


    def get_applications(self):
        with open(config_file, "r") as file:
            return [line.strip() for line in file if line.strip()]


    def read_sizes_from_file(self):
        sizes = []
        with open(window_sizes_file, "r") as f:
            for line in f:
                if line.strip():
                    size = line.strip().split(",")
                    if len(size) == 6:
                        try:
                            title, app, width, height, left, top = size
                            sizes.append((title, app, int(width), int(height), int(left), int(top)))
                        except ValueError:
                            print(f"Skipping line {line.strip()} as it contains non-integer values")
                    else:
                        print(f"Skipping line {line.strip()} as it does not have 6 values")
        return sizes


    def write_sizes_to_file(self, sizes):
        with open(window_sizes_file, "w") as f:
            for size in sizes:
                f.write(f"{size[0]},{size[1]},{size[2]},{size[3]},{size[4]},{size[5]}\n")


    def grab_window_sizes(self):
        sizes = self.get_window_sizes()
        self.write_sizes_to_file(sizes)


    def set_window_sizes(self):
        sizes = self.read_sizes_from_file()
        hwnds = []
        win32gui.EnumWindows(lambda hwnd, param: param.append(hwnds.append(hwnd)), [])
        for title, app, width, height, left, top in sizes:
            for hwnd in hwnds:
                if win32gui.IsWindowVisible(hwnd):
                    hwnd_title = win32gui.GetWindowText(hwnd)
                    if app.lower() in hwnd_title.lower():
                        try:
                            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, left, top, width, height, win32con.SWP_SHOWWINDOW)
                        except pywintypes.error as e:
                            with open(errorlog_file, "w") as f:
                                f.write(f"Error setting window size for '{title}': {e}")

        keep_hwnd = win32gui.FindWindow(None, "Keep Size")
        if keep_hwnd != 0:
            win32gui.SetWindowPos(keep_hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = KeepSize()
    window.show()
    sys.exit(app.exec_())
