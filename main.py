import os
import re
import sys
import threading
from datetime import datetime

import pdfplumber
from pandas import DataFrame
from PyQt6 import QtGui
from PyQt6.QtWidgets import QApplication, QDialog, QFileDialog
from PyQt6.uic import loadUi

basedir = os.path.dirname(__file__)


class WelcomeScreen(QDialog):
    def __init__(self):
        super(WelcomeScreen, self).__init__()
        loadUi(os.path.join(basedir, "welcome.ui"), self)
        self.btn_browse.clicked.connect(self.file_dailog)
        self.btn_submit.clicked.connect(self.process_thread)

    def file_dailog(self):
        self.clear_fields()
        filename, _ = QFileDialog.getOpenFileName(
            parent=self,
            caption="Select your CAMS PDF file",
            directory=os.path.expanduser("~"),
            filter="pdf(*.pdf)",
        )
        self.lbl_path.setText(filename)

    def csv_export(self, df):
        os.makedirs(os.path.join(os.path.expanduser("~"), "Downloads"), exist_ok=True)
        file_name = f"CAMS_Data_{datetime.now().strftime('%d_%m_%Y_%H_%M')}.csv"
        save_file_path = os.path.join(os.path.expanduser("~"), "Downloads", file_name)
        df.to_csv(save_file_path, index=False)
        self.lbl_message.setText(
            "Process completed, file saved in your Downloads folder"
        )

    def clean_txt(self, col):
        col.replace(r",", "", regex=True, inplace=True)
        col.replace(r"\(", "-", regex=True, inplace=True)
        col.replace(r"\)", " ", regex=True, inplace=True)
        return col

    def clear_fields(self):
        self.lbl_message.clear()
        self.lbl_path.clear()
        self.le_pwd.clear()
        self.le_pwd.setPlaceholderText("Document Password")
        self.le_pwd.setEnabled(True)

    def process_thread(self):
        self.lbl_message.setText("Processing please wait...")
        threading.Thread(target=self.file_processing).start()

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
                    # Converting text to list format
                    final_text = final_text.splitlines()
                final_df = self.extract_text(final_text)
                self.csv_export(final_df)

            except Exception as err_msg:
                if repr(err_msg) == "PdfminerException(PDFPasswordIncorrect())":
                    self.lbl_message.setText(
                        "File is Encrypted, please enter your password"
                    )
                    self.le_pwd.setEnabled(True)
                    self.le_pwd.setPlaceholderText("Document Password")
                else:
                    self.lbl_message.setText(repr(err_msg))
                    print(repr(err_msg))
        else:
            self.lbl_message.setText("Please select your CAMS PDF file...")

    def extract_text(self, doc_txt):
        # Defining RegEx patterns
        # Extracting Fund Name
        fund_name = re.compile(
            r"^([a-z0-9]{3,}+)-(.*?FUND.+)(\s\-\sISIN)", flags=re.IGNORECASE
        )

        # Extracting ISIN Number
        isin_num = re.compile(r"ISIN:\s*(\w+)", re.IGNORECASE)

        # Extracting Folio information
        folio_pat = re.compile(r"Folio No:\s*(\d+)", flags=re.IGNORECASE)

        # Extracting PAN information
        pan_pat = re.compile(r"PAN:\s*(\w{10})", flags=re.IGNORECASE)

        # Extracting Transaction data
        trans_details = re.compile(
            r"(^\d{2}-\w{3}-\d{4})(\s.+?\s(?=[\d(]))([\d\(]+[,\d.]+\d+[.\d\)]+)(\s[\d\(\,\.\)]+)(\s[\d\,\.]+)(\s[\d,\.]+)"
        )

        line_itms = []
        for idx, txt in enumerate(doc_txt):
            # Grab mutual fund name
            fund_chk = fund_name.match(txt)
            if fund_chk:
                fnd_name = fund_chk.group(2)

            # Grab Mutual fund ISIN number with index method
            isin_chk = isin_num.search(txt)
            if isin_chk:
                if len(isin_chk.group(1)) != 12:
                    next_match = re.match(r"^\w+", doc_txt[idx + 1])
                    isin = str(isin_chk.group(1)) + str(next_match.group(0))
                else:
                    isin = str(isin_chk.group(1))

            # Grab folio number & Investor name using folio pattern with index method
            folio_chk = folio_pat.search(txt)
            if folio_chk:
                folio = folio_chk.group(1)
                # Identiry current index position add 1 to grab the investor name from next line
                investor_name = doc_txt[idx + 1].title()

            # Grab Investor PAN number
            pan_check = pan_pat.search(txt)
            if pan_check:
                pan_num = pan_check.group(1)

            # Grab Mutual transaction details
            trn_txt = trans_details.search(txt)
            if trn_txt:
                date = trn_txt.group(1)
                description = trn_txt.group(2)
                amount = trn_txt.group(3)
                units = trn_txt.group(4)
                price = trn_txt.group(5)
                unit_bal = trn_txt.group(6)
                line_itms.append(
                    [
                        investor_name,
                        pan_num,
                        folio,
                        isin,
                        fnd_name,
                        date,
                        description,
                        amount,
                        units,
                        price,
                        unit_bal,
                    ],
                )

            df = DataFrame(
                line_itms,
                columns=[
                    "Investor_Name",
                    "PAN_Number",
                    "Folio",
                    "ISIN",
                    "Fund_Name",
                    "Date",
                    "Description",
                    "Amount",
                    "Units",
                    "Price",
                    "Unit_Balance",
                ],
            )

            for col in ["Amount", "Units", "Price", "Unit_Balance"]:
                self.clean_txt(df[col])
                df[col] = df[col].astype("float")
        return df


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
