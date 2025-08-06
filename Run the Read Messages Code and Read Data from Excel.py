import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QProgressBar, QLabel, QScrollArea
)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont
import runpy

def read_bins_from_excel():
    """Read all bins and their percentages from the Excel file."""
    try:
        df = pd.read_excel("latest_message.xlsx")
        bins = []
        for _, row in df.iterrows():
            address = str(row["Address"])
            percent = int(row["Message"])
            bins.append((address, percent))
        return bins
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []

class SmartGarbageApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Smart Garbage Monitoring System")
        self.showFullScreen()

        # Main style
        self.setStyleSheet("""
            QWidget {
                background: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:1,
                    stop:0 rgba(0, 0, 0, 255), stop:1 rgba(29, 107, 119, 255));
            }
            QScrollArea {
                border: none;
            }
        """)

        # Layouts
        self.main_layout = QVBoxLayout()
        self.main_layout.setAlignment(Qt.AlignTop)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)

        # Title
        title = QLabel("Smart Garbage Monitoring System")
        title.setFont(QFont("Arial", 40, QFont.Bold))
        title.setStyleSheet("""
            color: white; 
            background: transparent;
            padding: 20px;
            margin-bottom: 10px;
        """)
        self.main_layout.addWidget(title)

        # Scroll area for bins
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        # Bin container widget
        self.bin_widget = QWidget()
        self.bin_widget.setStyleSheet("background: transparent;")
        self.bin_layout = QHBoxLayout(self.bin_widget)
        self.bin_layout.setSpacing(15)
        self.bin_layout.setContentsMargins(15, 5, 15, 5)

        self.scroll_area.setWidget(self.bin_widget)
        self.main_layout.addWidget(self.scroll_area)
        self.setLayout(self.main_layout)

        # Dictionary to keep widgets per bin
        self.bin_widgets = {}

        # Timer to update every 10s
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_bins)
        self.timer.start(10000)

        self.update_bins()

    def update_bins(self):
        runpy.run_path("Read Messages and store in excel.py")
        bin_data = read_bins_from_excel()

        for address, percent in bin_data:
            if address not in self.bin_widgets:
                # Create new bin widget
                vbox = QVBoxLayout()
                vbox.setSpacing(5)
                
                # Address label
                label = QLabel(address)
                label.setFont(QFont("Arial", 20, QFont.Bold))
                label.setStyleSheet("""
                    color: white; 
                    background: transparent;
                    margin-bottom: 5px;
                """)
                label.setAlignment(Qt.AlignCenter)

                # Progress bar with gradient
                bar = QProgressBar()
                bar.setOrientation(Qt.Vertical)
                bar.setRange(0, 100)
                bar.setValue(percent)
                bar.setFixedSize(300, 550)
                bar.setFormat("%p%")
                bar.setAlignment(Qt.AlignCenter)
                bar.setFont(QFont("Arial", 15, QFont.Bold))
                bar.setStyleSheet(f"""
                    QProgressBar {{
                        color: white;
                        border: 3px solid #444;
                        border-radius: 12px;
                        background: #1a1a1a;
                        padding: 5px;
                        margin: 0 10px;
                    }}
                    QProgressBar::chunk {{
                        {self.get_gradient_style(percent)}
                        border-radius: 20px;
                    }}
                """)

                vbox.addWidget(label)
                vbox.addWidget(bar, alignment=Qt.AlignCenter)
                self.bin_layout.addLayout(vbox)
                self.bin_widgets[address] = bar
            else:
                # Update existing bar
                bar = self.bin_widgets[address]
                bar.setValue(percent)
                bar.setStyleSheet(f"""
                    QProgressBar {{
                        color:white;
                        border: 3px solid #444;
                        border-radius: 12px;
                        background: #1a1a1a;
                        padding: 5px;
                        margin: 0 10px;
                    }}
                    QProgressBar::chunk {{
                        {self.get_gradient_style(percent)}
                        border-radius: 20px;
                    }}
                """)

    def get_gradient_style(self, percent):
        if percent <= 50:
            return """
                background: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:1,
                    stop:0 rgba(0, 255, 0, 255), stop:1 rgba(0, 200, 0, 255));
            """
        elif percent < 90:
            return """
                background: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:1,
                    stop:0 rgba(255, 165, 0, 255), stop:1 rgba(255, 140, 0, 255));
            """
        else:
            return """
                background: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:1,
                    stop:0 rgba(255, 0, 0, 255), stop:1 rgba(200, 0, 0, 255));
            """

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SmartGarbageApp()
    window.show()
    sys.exit(app.exec_())