# CSV to PDF Converter

This Python script allows you to convert tab-delimited `.csv` files into clean, human-readable PDF documents. It is specifically designed to extract only the 3rd and 10th columns from the CSV and format them with proper spacing for readability.

## Features

- File picker dialog for selecting any `.csv` file
- Automatically extracts the 3rd and 10th columns
- Adds spacing between columns and rows for better layout
- Saves the output as a PDF in the same directory as the input file

## Requirements

Install the required Python packages:

```bash
pip install pandas pymupdf

## Usage

- Run the script.
- A file picker dialog will appear.
- Select the .csv file you want to convert.
- The script will generate a PDF with the suffix _filtered.pdf in the same folder.

## Notes

- The script assumes the CSV is encoded in UTF-16 and uses tab (\\t) as the delimiter.
- Only the 3rd and 10th columns (by position) are included in the PDF output.