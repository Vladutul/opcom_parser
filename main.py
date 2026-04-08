import sys
import pandas as pd
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QTextEdit, QLabel
from PyQt6.QtCore import QTimer, QTime

class OpcomApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        
        # Timer pentru verificarea orei (verifică la fiecare minut)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.check_schedule)
        self.timer.start(60000) 

    def initUI(self):
        self.setWindowTitle('OPCOM Automation Tool')
        layout = QVBoxLayout()

        self.label = QLabel('Status: Așteptare interval 13:00 - 16:00', self)
        layout.addWidget(self.label)

        self.log_area = QTextEdit(self)
        self.log_area.setReadOnly(True)
        layout.addWidget(self.log_area)

        self.btn_manual = QPushButton('Rulează Manual Acum', self)
        self.btn_manual.clicked.connect(self.run_process)
        layout.addWidget(self.btn_manual)

        self.setLayout(layout)
        self.resize(400, 300)

    def check_schedule(self):
        current_time = QTime.currentTime()
        start_time = QTime(13, 0)
        end_time = QTime(16, 0)

        if start_time <= current_time <= end_time:
            self.log_area.append("Ora potrivită detectată! Pornesc extragerea...")
            self.run_process()

    def run_process(self):
        # 1. Scraping (Exemplu conceptual)
        self.log_area.append("Extrag date de pe OPCOM...")
        
        # Aici vine logica de requests/beautifulsoup
        data = {'Ora': list(range(24)), 'Pret': [100 + i for i in range(24)]}
        df = pd.DataFrame(data)

        # 2. Export Excel
        filename = "date_opcom.xlsx"
        df.to_excel(filename, index=False)
        self.log_area.append(f"Date salvate în {filename}")

        # 3. Trimitere Email
        self.send_email(filename)

    def send_email(self, attachment):
        # Logica de SMTP (Gmail/Outlook)
        self.log_area.append("Email trimis cu succes către echipă!")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = OpcomApp()
    ex.show()
    sys.exit(app.exec())