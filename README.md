For Extracting URL From Hyper Links [Get URL From Hyperlink Excel Macro](https://www.notion.so/Get-URL-From-Hyperlink-Excel-Macro-761a6f13dc7040aa89aaea16420aa128?pvs=21) 

# Excel-Based File Downloader

This script downloads files from URLs specified in an Excel file and saves them to a specified directory. The filenames and download URLs are read from an Excel file, and the script handles downloading the files, tracking progress, and logging successes and failures.

## Prerequisites

- Python 3.x
- Required Python packages:
  - `openpyxl`
  - `wget`
  - `tqdm`
  - `pathlib`

## Installation

1. Clone this repository or download the script files.

2. Ensure you have Python installed. You can download it from [python.org](https://www.python.org/).

3. Install the required Python packages using pip:

   ```bash
   pip install openpyxl wget tqdm


Here's a comprehensive README file for your script. This README includes a description, prerequisites, installation steps, usage instructions, and additional notes for potential users.

markdown
Copy code
# Excel-Based File Downloader

This script downloads files from URLs specified in an Excel file and saves them to a specified directory. The filenames and download URLs are read from an Excel file, and the script handles downloading the files, tracking progress, and logging successes and failures.

## Prerequisites

- Python 3.x
- Required Python packages:
  - `openpyxl`
  - `wget`
  - `tqdm`
  - `pathlib`

## Installation

1. Clone this repository or download the script files.

2. Ensure you have Python installed. You can download it from [python.org](https://www.python.org/).

3. Install the required Python packages using pip:

   ```bash
   pip install openpyxl wget tqdm
Usage
Prepare your Excel file:

The Excel file should contain two columns:
Column A: Filenames for the downloaded files.
Column B: URLs from which to download the files.
Save your Excel file to a known location, for example: C:\\File\\File.xlsx.
Set up directories:

Ensure there is a directory to save the downloaded files, for example: C:\\File\\Downloads.
Run the script:

Update the paths in the script to match the locations of your Excel file and download directory.

Run the script using Python:

bash
Copy code
python download_files.py
The script will read the Excel file, download each file from the specified URLs, and save them with the specified filenames. Progress will be displayed in the terminal.

Script Details
The script performs the following steps:

Loads the Excel file specified by excel_file_path.
Reads filenames and download URLs from the Excel file.
Iterates over each row in the Excel file and attempts to download the file from the URL, saving it with the specified filename in the download_directory.
Tracks progress using tqdm and logs any download failures with reasons.
Outputs a summary of the total files processed, successful downloads, and failed downloads at the end.
Example Excel File
Filename	Download URL
file1.pdf	http://example.com/file1.pdf
file2.pdf	http://example.com/file2.pdf
...	...
Error Handling
The script checks for missing filenames or URLs and logs them as failures.
It handles exceptions during the download process and logs the reason for any failure.
Notes
Ensure that the URLs in the Excel file are accessible and the files are available for download.
Adjust the total_rows variable if your Excel file has more or fewer rows.
The script currently assumes that the Excel file has no headers. Adjust the script if your file includes headers.
License
This project is licensed under the MIT License. See the LICENSE file for details.

Acknowledgments
This script uses the openpyxl library for reading Excel files.
It uses the wget library for downloading files.
The tqdm library is used for displaying progress bars.
