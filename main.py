import sys, os
from PyQt5.uic import loadUi
from PyQt5 import QtGui
from PyQt5.QtWidgets import QDialog, QApplication, QFileDialog
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
        self.chk_password.toggled.connect(self.enable_pw_input)
        self.btn_submit.clicked.connect(self.process_thread)


    def file_dailog(self):
        self.clear_fields()
        filename, _ = QFileDialog.getOpenFileName(
            parent=self,
            caption="Select your CAMS PDF file",
            directory=os.getcwd(),
            filter="pdf(*.pdf)",
        )
        self.lbl_path.setText(filename)

    def enable_pw_input(self):
        if self.chk_password.isChecked() == True:
            self.le_pwd.setEnabled(True)
            self.le_pwd.setPlaceholderText("Document Password")
            self.lbl_message.clear()
        else:
            self.le_pwd.setEnabled(False)
            self.le_pwd.setPlaceholderText("")
            self.le_pwd.clear()

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

                self.extract_text(final_text)

            except:
                self.lbl_message.setText("Encrypted file, please enter your password")
        else:
            self.lbl_message.setText("Please select your CAMS PDF file..")


    def cleanup_fund_name(self, text):
        text = re.sub(r'Non-Demat', '', text)
        return text

    def extract_funds_details(self, text):
        # Regex patterns to extract the required details
        fund_name_pat = re.compile(r'^(.*?)\s*-\s*ISIN')
        isin_pat = re.compile(r'ISIN:\s*(\w+)')
        advisor_pat = re.compile(r'Advisor:\s*(\w+)')
        registrar_pat = re.compile(r'Registrar :\s*(\w+)')

        # Extract registrar
        registrar_match = registrar_pat.search(text)
        registrar = registrar_match.group(1) if registrar_match else None
        # lets remove registrar first so we can combine multiline text of other values below
        text = re.sub(r' Registrar :\s*(\w+)\s*', '', text)
        # Extract fund name
        fund_name_match = fund_name_pat.search(text)
        fund_name = None
        if fund_name_match:
            fund_name = fund_name_match.group()
            fund_name = re.sub(r'-\s*ISIN', '', fund_name) # remove the -ISIN at end of the line
            fund_name = re.sub(r'\(\s*Non-Demat\s*\)', '', fund_name) # remove the non-demat message
            fund_name = re.sub(r'\(formerly.*\)', '', fund_name) # remove older names for the fund
            fund_name = re.sub(r'\(erstwhile.*\)', '', fund_name) # remove older names for the fund
            fund_name = re.sub(r'^.*?-', '', fund_name) # remove the 4-5 letter code for the fund at start of fund name
        # fund_name = cleanup_fund_name(fund_name)
        # Extract ISIN
        isin_match = isin_pat.search(text)
        isin = isin_match.group(1) if isin_match else None

        # Extract advisor
        advisor_match = advisor_pat.search(text)
        advisor = advisor_match.group(1) if advisor_match else None

        return fund_name, isin, advisor, registrar


    # sometimes text overflows , so depend on the order of text lines as well
    # format is as follows
    # Folio No:XX      PAN:XX      KYC:OK PAN:OK
    # USER_ACCOUNT_NAME
    # FUND_NAME: YYYYY YYY YYY YYY           Registrar: UUUU
    # FUND_NAME_CONTINUED (optional)
    # Nominee 1:
    def extract_text(self, doc_txt):
        # Defining RegEx patterns
        folio_pat = re.compile(r"^Folio No:\s*\d+ / \d+", flags=re.IGNORECASE)
        fund_name_pat = re.compile(r".*Fund.*ISIN.*", flags=re.IGNORECASE)
        nominee_search = re.compile(r"^Nominee", flags=re.IGNORECASE)
        trans_details = re.compile(
            r"(^\d{2}-\w{3}-\d{4})(\s.+?\s(?=[\d(]))([\d\(]+[,.]\d+[.\d\)]+)(\s[\d\(\,\.\)]+)(\s[\d\,\.]+)(\s[\d,\.]+)"
        )  # Extracting Transaction data
        line_count_after_folio = 99; # start with high value as we want to init it when we find folio
        line_itms = []
        for i in doc_txt.splitlines():
            line_count_after_folio += 1
            # first check if this is folio line
            folio_match = folio_pat.search(i)
            if folio_match:
                folio = folio_match.group()
                folio = folio[10:]
                line_count_after_folio = 0

            fund_match = fund_name_pat.search(i)
            if fund_match:
                fund_name_line = i
                # we will process this only in the next line to see if any more text is seen
                continue
            elif line_count_after_folio == 2:
                fund_name_line = i
                continue

                # fun_name, isin, advisor, registrar = self.extract_funds_details(i)
            if line_count_after_folio == 3:
                # if this line did not contain nominee, that means we are continuing fund name
                nominee_match = nominee_search.search(i)
                if nominee_match is None:
                    fund_name_line += " " + i
                fun_name, isin, advisor, registrar = self.extract_funds_details(fund_name_line)

            txt = trans_details.search(i)
            if txt:
                date = txt.group(1)
                description = txt.group(2)
                amount = txt.group(3)
                units = txt.group(4)
                price = txt.group(5)
                unit_bal = txt.group(6)
                line_itms.append(
                    [fun_name, folio, date, units, price, unit_bal,amount, description, isin, advisor, registrar]
                )

            df = DataFrame(
                line_itms,
                columns=[
                    "Fund_name",
                    "Folio",
                    "Date",
                    "Units",
                    "Price",
                    "Unit_balance",
                    "Amount",
                    "Description",
                    "ISIN",
                    "Advisor",
                    "Registrar",
                ],
            )
            self.clean_txt(df.Amount)
            self.clean_txt(df.Units)
            self.clean_txt(df.Price)
            self.clean_txt(df.Unit_balance)

            df.Amount = df.Amount.astype("float")
            df.Units = df.Units.astype("float")
            df.Price = df.Price.astype("float")
            df.Unit_balance = df.Unit_balance.astype("float")

            file_name = f'CAMS_data_{datetime.now().strftime("%d_%m_%Y_%H_%M")}.csv'
            save_file = os.path.join(os.path.expanduser("~"), "Downloads", file_name)

            try:
                df.to_csv(save_file, index=False)
                self.lbl_message.setText(
                    "Process completed, file saved in Downloads folder"
                )

            except Exception as e:
                self.lbl_message.setText(e)

    def clean_txt(self, x):
        x.replace(r",", "", regex=True, inplace=True)
        x.replace(r"\(", "-", regex=True, inplace=True)
        x.replace(r"\)", " ", regex=True, inplace=True)
        return x

    def clear_fields(self):
        self.lbl_message.clear()
        self.lbl_path.clear()
        self.chk_password.setChecked(False)

    def process_thread(self):
        self.lbl_message.setText("Processing please wait...")
        threading.Thread(target=self.file_processing).start()


# Main
app = QApplication(sys.argv)
widget = WelcomeScreen()
widget.setWindowTitle(" ")
widget.setWindowIcon(QtGui.QIcon(os.path.join(basedir, "icons", "app_icon.svg")))
widget.show()

try:
    sys.exit(app.exec_())

except:
    print("exiting")
