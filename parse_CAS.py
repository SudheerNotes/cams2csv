import sys
import re
import pdfplumber
import pandas as pd

# Open the PDF file
def openfile(pdf_file, password=None):
    global pdf
    try:
        pdf = pdfplumber.open(pdf_file, password=password)
    except Exception as e:
        print(f"Error opening the file: {str(e)}")
        sys.exit(1)

# Get the total number of pages in the PDF
def get_page_count():
    return len(pdf.pages)

# Extract tables from each page
def get_tables():
    tables = []
    for page in pdf.pages:
        try:
            tables_per_page = page.find_tables()
            tables += tables_per_page
        except Exception as e:
            print(f"Error extracting tables on page {pdf.pages.index(page) + 1}: {str(e)}")
            continue
    return tables

# Function to filter out unwanted rows based on the first element of the row
def remove_unwanted_rows(data):
    unwanted_strings = ["Mutual Fund Folios (F)", "Equities (E)", "Equity Shares"]
    i = 0
    for i in range(len(data)):
        if data[i][0].startswith("Summary of value"):
            i += 2  # Skip the next 2 rows
            break
        if any(data[i][0].startswith(unwanted) for unwanted in unwanted_strings):
            continue
        break
    return data[i:]

# Create a new DataFrame structure for transactions
def create_new_transaction_table():
    columns = [
        "Date", "Fund", "Folio", "Transaction Details",
        "Amount", "Stamp Duty", "NAV", "Units", "Opening Balance", "Closing Balance"
    ]
    return pd.DataFrame(columns=columns)

# Function to transform a table into a transaction row
def transform_transactions_table(df, source):
    try:
        new_row = {
            "Date": source[2][0],
            "Fund": source[0][0],
            "Folio": source[0][5] if len(source[0]) > 5 else None,
            "Transaction Details": source[2][1],
            "Amount": source[2][2],
            "Stamp Duty": source[2][3],
            "NAV": source[2][4],
            "Units": source[2][6] if len(source[2]) > 6 else None,
            "Opening Balance": source[1][5] if len(source[1]) > 5 else None,
            "Closing Balance": source[3][5] if len(source[3]) > 5 else None
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    except IndexError as e:
        print(f"Error processing transaction row: {str(e)}")
    return df

# Clean a cell by removing unwanted characters and commas
def clean_cell(cell):
    if cell is None:
        return cell
    cell = cell.replace("`", "").replace("\n", " ")
    cell = re.sub(r'(?<=\d),(?=\d)', '', cell)  # Remove commas between digits
    if not re.search(r'[a-zA-Z]', cell):  # Check if there are no alphabets in the string
        return cell.replace(' ', '')  # Replace commas with an empty string
    return cell.strip() 

# Convert columns with numeric values to numeric data type
def convert_numeric_columns(df):
    for col in df.columns:
        try:
            df[col] = pd.to_numeric(df[col])
        except (ValueError, TypeError):
            # If conversion fails, the column remains unchanged
            pass
    return df

# Function to merge adjacent tables, handling transactions
def merge_adjacent_tables(tables):
    merged_tables = []
    current_batch = None
    transactions_parsing = False

    for table in tables:
        data = table.extract()

        # Skip tables with only a header row or empty data
        if len(data) <= 1:
            continue

        if transactions_parsing or data[0][0] == 'Date':
            if data[0][0] == 'Date':  # Start new transaction parsing
                transactions_parsing = True
                df = create_new_transaction_table()
                df = transform_transactions_table(df, data[1:])
                continue
            elif len(data) > 3 and data[1][1] == 'Opening Balance' and data[3][1] == 'Closing Balance':
                df = transform_transactions_table(df, data)
                continue
            else:
                transactions_parsing = False
                df = df.applymap(lambda x: clean_cell(x) if isinstance(x, str) else x)
                df = convert_numeric_columns(df)
                merged_tables.append(df)
                continue

        # Remove unwanted rows
        data = remove_unwanted_rows(data)

        # Convert table to DataFrame
        try:
            df = pd.DataFrame(data[1:], columns=data[0])
        except Exception as e:
            print(f"Error converting table to DataFrame: {str(e)}")
            continue

        # Clean columns and cells
        df.columns = [clean_cell(col) for col in df.columns]
        df = df.applymap(lambda x: clean_cell(x) if isinstance(x, str) else x)

        # Handle ISIN UCC split if the column exists
        if "ISIN UCC" in df.columns:
            isin_ucc_split = df["ISIN UCC"].str.split(" ", n=1, expand=True)
            df["ISIN"] = isin_ucc_split[0]
            df["UCC"] = isin_ucc_split[1]
            df.drop(columns=["ISIN UCC"], inplace=True)

        # Convert numeric columns
        df = convert_numeric_columns(df)

        # Merge or append tables based on column match
        if current_batch is None:
            current_batch = df
        elif current_batch.columns.equals(df.columns):
            current_batch = pd.concat([current_batch, df], ignore_index=True)
        else:
            merged_tables.append(current_batch)
            current_batch = df

    if current_batch is not None:
        merged_tables.append(current_batch)

    return merged_tables

# Save merged tables to an Excel file
def save_tables_to_excel(tables, excel_file="extracted_tables.xlsx"):
    merged_tables = merge_adjacent_tables(tables)

    # Write merged tables to Excel file, each in a different sheet
    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        for idx, df in enumerate(merged_tables):
            sheet_name = f"Table_{idx + 1}"
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print(f"Tables saved to {excel_file}")

# Main function
if __name__ == "__main__":
    password = "BKCPS2524L"
    pdf_file = "C:\\Users\\Siva\\Downloads\\NSDLe-CAS_112231234_AUG_2024.PDF"

    try:
        openfile(pdf_file, password=password)
        page_count = get_page_count()
        print(f"Page count of {pdf_file}: {page_count}")
        tables = get_tables()
        save_tables_to_excel(tables, "extracted_tables.xlsx")
        print(f"Extracted {len(tables)} tables")
        pdf.close()
    except FileNotFoundError:
        print(f"Error: File '{pdf_file}' not found.")
        sys.exit(1)
    except pdfplumber.pdfminer.pdfdocument.PDFPasswordIncorrect:
        print(f"Please check the password: {password} to open the file.")
        sys.exit(1)
 