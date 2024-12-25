import sys, os
from PyQt6.uic import loadUi
from PyQt6 import QtGui
from PyQt6.QtWidgets import QDialog, QApplication, QFileDialog
import pdfplumber
import re
from pandas import DataFrame
from datetime import datetime
import threading


basedir = os.path.dirname(__file__)

class WelcomeScreen(QDialog):
    def __init__(self):
        super(WelcomeScreen, self).__init__()
        loadUi(os.path.join(basedir,"welcome.ui"), self)
        self.btn_browse.clicked.connect(self.file_dailog)
        self.btn_submit.clicked.connect(self.process_thread)

    def file_dailog(self):
        self.clear_fields()
        self.le_pwd.setPlaceholderText("Document Password")
        self.le_pwd.setEnabled(True)
        filename, _ = QFileDialog.getOpenFileName(
            parent=self,
            caption="Select your CAMS PDF file",
            directory=os.getcwd(),
            filter="pdf(*.pdf)")
        self.lbl_path.setText(filename)

    def file_processing(self):
        file_path = self.lbl_path.text()
        doc_pwd = self.le_pwd.text()
        final_text = ""

        if not len(file_path) == 0:
            try:
                with pdfplumber.open(file_path, password=doc_pwd) as pdf:
                    for i in range(len(pdf.pages)):
                        txt = pdf.pages[i].extract_text()
                        final_text = final_text + "\n" + txt
                    pdf.close()
                final_df = self.extract_text(final_text)
                self.csv_export(final_df)

            except Exception as err_msg:
                if repr(err_msg) == "PDFPasswordIncorrect()":
                    self.lbl_message.setText("File is Encrypted, please enter your password")
                    self.le_pwd.setEnabled(True)
                    self.le_pwd.setPlaceholderText("Document Password")
                else:
                    self.lbl_message.setText(repr(err_msg))
        else:
            self.lbl_message.setText("Please select your CAMS PDF file..")

    def extract_text(self, doc_txt):

        # Defining RegEx patterns
        folio_pat = re.compile(r"(^Folio No:\s\d+)(?:\s.*)", flags=re.IGNORECASE)  # Extracting Folio information
        fund_name = re.compile(r"^([a-z0-9]{3,}+)-(.*?FUND)", flags=re.IGNORECASE) # Extracting Fund Name
        trans_details = re.compile(r"(^\d{2}-\w{3}-\d{4})(\s.+?\s(?=[\d(]))([\d\(]+[,.]\d+[.\d\)]+)(\s[\d\(\,\.\)]+)(\s[\d\,\.]+)(\s[\d,\.]+)")  # Extracting Transaction data

        line_itms = []
        for i in doc_txt.splitlines():

            fund_chk = fund_name.match(i)
            if fund_chk:    
                fun_name = fund_chk.group(0)

            folio_chk = folio_pat.match(i)
            if folio_chk:
                folio = folio_chk.group(1)

            txt = trans_details.search(i)
            if txt:
                date = txt.group(1)
                description = txt.group(2)
                amount = txt.group(3)
                units = txt.group(4)
                price = txt.group(5)
                unit_bal = txt.group(6)
                line_itms.append([folio, fun_name, date, description, amount, units, price, unit_bal])

            df = DataFrame(
                line_itms,
                columns=[
                    "Folio",
                    "Fund_name",
                    "Date",
                    "Description",
                    "Amount",
                    "Units",
                    "Price",
                    "Unit_balance",
                ])
            self.clean_txt(df.Amount)
            self.clean_txt(df.Units)
            self.clean_txt(df.Price)
            self.clean_txt(df.Unit_balance)

            df.Amount = df.Amount.astype("float")
            df.Units = df.Units.astype("float")
            df.Price = df.Price.astype("float")
            df.Unit_balance = df.Unit_balance.astype("float")
        return df 

    def csv_export(self, df):
        file_name = f'CAMS_data_{datetime.now().strftime("%d_%m_%Y_%H_%M")}.csv'
        save_file = os.path.join(os.path.expanduser("~"), "Downloads", file_name)
        os.makedirs(os.path.join(os.path.expanduser("~"), "Downloads"), exist_ok=True)
        df.to_csv(save_file, index=False)
        self.lbl_message.setText("Process completed, file saved in your Downloads folder")  

    def clean_txt(self, x):
        x.replace(r",", "", regex=True, inplace=True)
        x.replace(r"\(", "-", regex=True, inplace=True)
        x.replace(r"\)", " ", regex=True, inplace=True)
        return x

    def clear_fields(self):
        self.lbl_message.clear()
        self.lbl_path.clear()
        self.le_pwd.clear()
        self.le_pwd.setPlaceholderText("")
        self.le_pwd.setEnabled(False)

    def process_thread(self):
        self.lbl_message.setText("Processing please wait...")
        threading.Thread(target=self.file_processing).start()

# Main
app = QApplication(sys.argv)
window = WelcomeScreen()
window.setWindowTitle(" ")
window.setWindowIcon(QtGui.QIcon(os.path.join(basedir, "icons", "app_icon.svg")))
window.show()

try:
    sys.exit(app.exec())

except:
    print("exiting")
