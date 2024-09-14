import sys
import os
from PyQt5.uic import loadUi
from PyQt5 import QtGui
from PyQt5.QtWidgets import QDialog, QApplication, QFileDialog, QMessageBox
from PyQt5.QtCore import pyqtSignal, QObject
import pdfplumber
import re
from pandas import DataFrame
from datetime import datetime
import threading
import time

basedir = os.path.dirname(__file__)

class Worker(QObject):
    error_occurred = pyqtSignal(str)
    processing_done = pyqtSignal(str)

    def __init__(self, file_path, doc_pwd):
        super().__init__()
        self.file_path = file_path
        self.doc_pwd = doc_pwd

    def run(self):
        try:
            self.file_processing()
        except Exception as e:
            self.error_occurred.emit(str(e))

    def file_processing(self):
        if not self.file_path:
            raise ValueError("Please select your CAMS PDF file.")

        final_text = ""
        try:
            with pdfplumber.open(self.file_path, password=self.doc_pwd) as pdf:
                final_text = "\n".join(page.extract_text() for page in pdf.pages)

            self.extract_text(final_text)

        except pdfplumber.pdfminer.pdfdocument.PDFPasswordIncorrect:
            raise ValueError("Encrypted file, please enter your password.")
        except Exception as e:
            raise ValueError(f"An error occurred: {e}")

    def extract_text(self, doc_txt):
        # Your extraction logic here...
        # At the end of processing, emit the success signal
        self.processing_done.emit("Process completed, file saved in Downloads folder")

class WelcomeScreen(QDialog):
    def __init__(self):
        super().__init__()
        loadUi(os.path.join(basedir, "welcome.ui"), self)
        self.btn_browse.clicked.connect(self.file_dialog)
        self.chk_password.toggled.connect(self.enable_pw_input)
        self.btn_submit.clicked.connect(self.process_thread)
        self.thread = None
        self.worker = None

    def file_dialog(self):
        self.clear_fields()
        filename, _ = QFileDialog.getOpenFileName(
            parent=self,
            caption="Select your CAMS PDF file",
            directory=os.getcwd(),
            filter="PDF files (*.pdf)",
        )
        self.lbl_path.setText(filename)

    def enable_pw_input(self):
        self.le_pwd.setEnabled(self.chk_password.isChecked())
        self.le_pwd.setPlaceholderText("Document Password" if self.chk_password.isChecked() else "")
        if not self.chk_password.isChecked():
            self.le_pwd.clear()
        self.lbl_message.clear()

    def show_error_message(self, message):
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Critical)
        msg_box.setText("Error")
        msg_box.setInformativeText(message)
        msg_box.setWindowTitle("Error")
        msg_box.exec_()

    def process_thread(self):
        self.lbl_message.setText("Processing, please wait...")
        file_path = self.lbl_path.text()
        doc_pwd = self.le_pwd.text()
        self.thread = threading.Thread(target=self.start_worker, args=(file_path, doc_pwd))
        self.thread.start()

    def start_worker(self, file_path, doc_pwd):
        self.worker = Worker(file_path, doc_pwd)
        self.worker.error_occurred.connect(self.show_error_message)
        self.worker.processing_done.connect(self.update_message)
        self.worker.run()

    def update_message(self, message):
        self.lbl_message.setText(message)

    def clear_fields(self):
        self.lbl_message.clear()
        self.lbl_path.clear()
        self.chk_password.setChecked(False)

    def __init__(self):
        super().__init__()
        loadUi(os.path.join(basedir, "welcome.ui"), self)
        self.btn_browse.clicked.connect(self.file_dialog)
        self.chk_password.toggled.connect(self.enable_pw_input)
        self.btn_submit.clicked.connect(self.process_thread)
        self.thread = None
        self.thread_running = False

    def file_dialog(self):
        self.clear_fields()
        filename, _ = QFileDialog.getOpenFileName(
            parent=self,
            caption="Select your CAMS PDF file",
            directory=os.getcwd(),
            filter="PDF files (*.pdf)",
        )
        self.lbl_path.setText(filename)

    def enable_pw_input(self):
        self.le_pwd.setEnabled(self.chk_password.isChecked())
        self.le_pwd.setPlaceholderText("Document Password" if self.chk_password.isChecked() else "")
        if not self.chk_password.isChecked():
            self.le_pwd.clear()
        self.lbl_message.clear()

    def show_error_message(self, message):
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Critical)
        msg_box.setText("Error")
        msg_box.setInformativeText(message)
        msg_box.setWindowTitle("Error")
        msg_box.exec_()

    def check_password_file(self):
        file_path = self.lbl_path.text()
        doc_pwd = self.le_pwd.text()
        try:
            with pdfplumber.open(file_path, password=doc_pwd) as pdf:
                return True
        except pdfplumber.pdfminer.pdfdocument.PDFPasswordIncorrect:
            return False
        except Exception as e:
            self.show_error_message(f"An error occurred: {e}")
        return False


    def file_processing(self):
        file_path = self.lbl_path.text()
        doc_pwd = self.le_pwd.text()
        if not file_path:
            self.show_error_message("Please select your CAMS PDF file.")
            return

        if False == self.check_password_file():
            self.show_error_message("Encrypted file, please enter your password.")
            return

        final_text = ""
        try:
            with pdfplumber.open(file_path, password=doc_pwd) as pdf:
                final_text = "\n".join(page.extract_text() for page in pdf.pages)

            self.extract_text(final_text)
        except Exception as e:
            self.show_error_message(f"An error occurred: {e}")

    def extract_funds_details(self, text):
        fund_name_pat = re.compile(r'^(.*?)\s*-\s*ISIN')
        isin_pat = re.compile(r'ISIN:\s*(\w+)')
        advisor_pat = re.compile(r'Advisor:\s*(\w+)')
        registrar_pat = re.compile(r'Registrar :\s*(\w+)')

        registrar_match = registrar_pat.search(text)
        registrar = registrar_match.group(1) if registrar_match else None

        text = re.sub(r' Registrar :\s*(\w+)\s*', '', text)

        fund_name_match = fund_name_pat.search(text)
        fund_name = fund_name_match.group(1) if fund_name_match else None
        fund_name = re.sub(r'(\s*-\s*ISIN|\s*\(formerly.*\)|\s*\(erstwhile.*\))', '', fund_name)
        fund_name = re.sub(r'^\w*-', '', fund_name)

        isin_match = isin_pat.search(text)
        isin = isin_match.group(1) if isin_match else None

        advisor_match = advisor_pat.search(text)
        advisor = advisor_match.group(1) if advisor_match else None

        return fund_name, isin, advisor, registrar

    def extract_text(self, doc_txt):
        folio_pat = re.compile(r"^Folio No:\s*\d+ / \d+", re.IGNORECASE)
        fund_name_pat = re.compile(r".*Fund.*ISIN.*", re.IGNORECASE)
        nominee_search = re.compile(r"^Nominee", re.IGNORECASE)
        trans_details = re.compile(
            r"(^\d{2}-\w{3}-\d{4})(\s.+?\s(?=[\d(]))([\d\(]+[,.]\d+[.\d\)]+)(\s[\d\(\,\.\)]+)(\s[\d\,\.]+)(\s[\d,\.]+)"
        )

        line_itms = []
        folio = ""
        fund_name_line = ""
        line_count_after_folio = 99

        for line in doc_txt.splitlines():
            line_count_after_folio += 1

            folio_match = folio_pat.search(line)
            if folio_match:
                folio = folio_match.group()[10:]
                line_count_after_folio = 0
                continue

            if fund_name_pat.search(line):
                fund_name_line = line
                continue

            if line_count_after_folio == 2:
                fund_name_line = line
                continue

            if line_count_after_folio == 3:
                if not nominee_search.search(line):
                    fund_name_line += " " + line

                fun_name, isin, advisor, registrar = self.extract_funds_details(fund_name_line)

            trans_match = trans_details.search(line)
            if trans_match:
                date, description, amount, units, price, unit_bal = trans_match.groups()
                line_itms.append(
                    [fun_name, folio, date, units, price, unit_bal, amount, description, isin, advisor, registrar]
                )

        df = DataFrame(
            line_itms,
            columns=[
                "Fund_name", "Folio", "Date", "Units", "Price", "Unit_balance",
                "Amount", "Description", "ISIN", "Advisor", "Registrar",
            ],
        )

        for col in ["Amount", "Units", "Price", "Unit_balance"]:
            df[col] = df[col].str.replace(",", "").str.replace("(", "-").str.replace(")", "").astype(float)

        file_name = f'CAMS_data_{datetime.now().strftime("%d_%m_%Y_%H_%M")}.csv'
        save_file = os.path.join(os.path.expanduser("~"), "Downloads", file_name)

        try:
            df.to_csv(save_file, index=False)
            self.lbl_message.setText("Process completed, file saved in Downloads folder")
        except Exception as e:
            self.show_error_message(f"Failed to save file: {e}")

    def clear_fields(self):
        self.lbl_message.clear()
        self.lbl_path.clear()
        self.chk_password.setChecked(False)

    def run_file_processing(self):
        while True:
            self.file_processing()
            time.sleep(1)

    def process_thread(self):
        self.lbl_message.setText("Processing, please wait...")
        if not self.thread or not self.thread.is_alive():
            self.thread = threading.Thread(target=self.run_file_processing)
            self.thread_running = True
            self.thread.start()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    widget = WelcomeScreen()
    widget.setWindowTitle("CAMS PDF Extractor")
    widget.setWindowIcon(QtGui.QIcon(os.path.join(basedir, "icons", "app_icon.svg")))
    widget.show()

    try:
        sys.exit(app.exec_())
    except SystemExit:
        print("Exiting")
    except Exception as e:
        print("unhandled error")
