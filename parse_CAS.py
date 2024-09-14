import sys
import re
import pdfplumber
import pandas as pd

# Open the PDF file
def openfile(pdf_file, password=None):
    global pdf
    pdf = pdfplumber.open(pdf_file, password=password)

# Get the total number of pages in the PDF
def get_page_count():
    return len(pdf.pages)

# Extract tables from each page
def get_tables():
    tables = []
    table_settings = {
        "vertical_strategy": "text",
        "horizontal_strategy": "text"
    }
    for i in range(len(pdf.pages)):
        tables_per_page = pdf.pages[i].find_tables()
        tables += tables_per_page
    return tables

# Function to filter out unwanted rows based on the first element of the row
def remove_unwanted_rows(data):
    unwanted_strings = ["Mutual Fund Folios (F)", "Equities (E)", "Equity Shares"]
    i = 0
    for i in range (0, len(data)):
        if data and data[i][0].startswith("Summary of value"):
            i=i+2
            break
        # Check if the first row exists and its first column matches any of the unwanted strings
        if data and any(data[i][0].startswith(unwanted) for unwanted in unwanted_strings):
            # Remove the first row and return the remaining data
            continue
        break
    # If no match, return the data as is
    return data[i:]

def create_new_transaction_table():
    columns = [
        "Date",
        "Fund",
        "Folio",
        "Transaction Details",
        "Amount",
        "Stamp Duty",
        "NAV",
        "Units",
        "Opening Balance",
        "Closing Balance"
    ]
    # Create an empty DataFrame with these columns
    df = pd.DataFrame(columns=columns)
    return df

def transform_transactions_table(df, source):
    new_row = {
        "Date": source[2][0],
        "Fund": source[0][0],
        "Folio": source[0][5],
        "Transaction Details": source[2][1],
        "Amount": source[2][2],
        "Stamp Duty": source[2][3],
        "NAV": source[2][4],
        "Units": source[2][6],
        "Opening Balance": source[1][5],
        "Closing Balance": source[3][5]
    }
    # Convert the new row to a DataFrame and use pd.concat() to add it
    new_row_df = pd.DataFrame([new_row])  # Create a DataFrame from the row dictionary
    df = pd.concat([df, new_row_df], ignore_index=True)
    return df

def merge_adjacent_tables(tables):
    merged_tables = []
    current_batch = None
    transactions_parsing = False

    for i in range(0, len(tables)):
        # Extract data from the table
        data = tables[i].extract()

        # Skip tables with only a header row
        if len(data) == 1:
            continue
        # Remove unwanted rows based on the first element of the row
        if transactions_parsing or data[0][0] == 'Date':
            if data[0][0] == 'Date':
                transactions_parsing = True
                df = create_new_transaction_table()
                df = transform_transactions_table(df, data[1:])
                continue
            elif data[1][1] == 'Opening Balance' and data[3][1] == 'Closing Balance':
                # continue transactions being parsed and added into a new table
                df = transform_transactions_table(df, data)
                continue
            else:
                transactions_parsing = False
                df = convert_numeric_columns(df)
                merged_tables.append(df)

        data = remove_unwanted_rows(data)

        # Convert the current table to a DataFrame
        df = pd.DataFrame(data[1:], columns=data[0])

        # Clean column names
        df.columns = [clean_cell(col) for col in df.columns]
        # Remove all occurrences of the backtick character "`" and remove commas between numbers in each cell
        df = df.applymap(lambda x: clean_cell(x) if isinstance(x, str) else x)

        # Convert columns with only numeric values to numeric data type
        df = convert_numeric_columns(df)

        # Check if "ISIN UCC" column exists and split it
        if "ISIN UCC" in df.columns:
            # Split the column into two columns based on space
            isin_ucc_split = df["ISIN UCC"].str.split(" ", n=1, expand=True)
            df["ISIN"] = isin_ucc_split[0]  # First part goes into "ISIN"
            df["UCC"] = isin_ucc_split[1]  # Second part goes into "UCC"
            df = df.drop(columns=["ISIN UCC"])  # Drop the original "ISIN UCC" column

        # If this is the first table in the batch, initialize the current_batch
        if current_batch is None:
            current_batch = df
        else:
            # If columns are the same as the previous table, merge them
            if current_batch.columns.equals(df.columns):
                current_batch = pd.concat([current_batch, df], ignore_index=True)
            else:
                # Print the columns that don't match
                print("Columns that don't match:")
                print("In current_batch but not in df:", current_batch.columns.difference(df.columns).tolist())
                print("In df but not in current_batch:", df.columns.difference(current_batch.columns).tolist())
                
                # If columns are different, save the current batch and start a new one
                merged_tables.append(current_batch)
                current_batch = df

    # Don't forget to add the last batch
    if current_batch is not None:
        merged_tables.append(current_batch)

    return merged_tables

# Function to clean a cell: remove backticks and remove commas between numbers
def clean_cell(cell):

    if cell is None:
        return
    # Remove backticks
    cell = cell.replace("`", "")
    cell = cell.replace("\n", " ")
    # Use regular expressions to remove commas between numbers
    cell = re.sub(r'(?<=\d),(?=\d)', '', cell)  # Remove commas between digits
    if not re.search(r'[a-zA-Z]', cell):  # Check if there are no alphabets in the string
        return cell.replace(' ', '')  # Replace commas with an empty string
    cell = cell.strip()
    return cell

# Function to convert columns with only numeric values to numeric data type
def convert_numeric_columns(df):
    for col in df.columns:
        try:
            df[col] = pd.to_numeric(df[col])
        except (ValueError, TypeError):
            # If conversion fails, the column remains unchanged
            pass
    return df
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
    # if len(sys.argv) < 2:
    #    print("Usage: python pdf_page_count.py <pdf_file> [password]")
    #    sys.exit(1)
    
    # pdf_file = sys.argv[1]
    password = "BKCPS2524L"
    pdf_file = "C:\\Users\\Siva\\Downloads\\NSDLe-CAS_112231234_AUG_2024.PDF"
    try:
        openfile(pdf_file, password=password)
        page_count = get_page_count()
        print(f"Page count of {pdf_file}: {page_count}")
        tables = get_tables()
        save_tables_to_excel(tables, "extracted_tables.xlsx")
        print(f"Got {len(tables)} tables")
        pdf.close()
    except FileNotFoundError:
        print(f"Error: File '{pdf_file}' not found.")
        sys.exit(1)
    except pdfplumber.pdfminer.pdfdocument.PDFPasswordIncorrect:
        print(f"Please recheck the password: {password} to open the file")
        sys.exit(1)
