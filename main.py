import openpyxl
import wget
import os

# Path to the Excel file
excel_path = "C:\\LRS\\LRS.xlsx"
# Directory to save the downloaded files
download_dir = "C:\\LRS\\Downloads"

# Ensure the download directory exists
os.makedirs(download_dir, exist_ok=True)

# Load the workbook and select the active sheet
wb_obj = openpyxl.load_workbook(excel_path, read_only=True)
sheet_obj = wb_obj.active

# Iterate over the rows
for row in sheet_obj.iter_rows(min_row=2, max_col=2):
    filename = sheet_obj.cell(row=row[0].row, column=1).value
    download_url = sheet_obj.cell(row=row[0].row, column=2).value

    # Check if the filename and URL are not empty
    if filename and download_url:
        try:
            # Define the output file path
            output_filepath = os.path.join(download_dir, filename)

            # Download the file
            wget.download(download_url, out=output_filepath)
            print(f"Downloaded {filename}")
        except Exception as e:
            print(f"Error downloading {filename}: {e}")
    else:
        print(f"Skipping row {row[0].row} due to missing filename or URL")

# Close the workbook
wb_obj.close()
