import glob
import os

import fitz  # PyMuPDF
import pandas as pd

folder_path = os.path.join(os.path.expanduser("~"), "Downloads")
csv_files = glob.glob(os.path.join(folder_path, "*.csv"))

pdf = fitz.open()
page_width, page_height = fitz.paper_size("a4")
margin = 50
line_height = 18
column_spacing = 150
max_lines_per_page = int((page_height - 2 * margin) / line_height)

current_page = None
current_line = 0
pages_added = 0


def ensure_page():
    global current_page, current_line, pages_added
    if current_page is None or current_line >= max_lines_per_page:
        current_page = pdf.new_page(width=page_width, height=page_height)
        current_line = 0
        pages_added += 1


def write_row(col1, col2):
    global current_line
    ensure_page()
    y = margin + current_line * line_height
    current_page.insert_text((margin, y), str(col1), fontsize=10, fontname="helv")
    current_page.insert_text(
        (margin + column_spacing, y), str(col2), fontsize=10, fontname="helv"
    )
    current_line += 1


for csv_file in csv_files:
    try:
        df = pd.read_csv(csv_file, delimiter="\t", encoding="utf-16", skiprows=1)
        if df.shape[1] < 10:
            print(f"Skipping {csv_file}: not enough columns.")
            continue

        df_filtered = df.iloc[:, [2, 9]]
        write_row(*df_filtered.columns)  # Write header once per file

        for _, row in df_filtered.iterrows():
            write_row(row.iloc[0], row.iloc[1])

        # os.remove(csv_file) # Keep commented out while development cycle to keep the csv files

    except Exception as e:
        print(f"Failed to process {csv_file}: {e}")

if pages_added > 0:
    output_pdf = os.path.join(folder_path, "Report.pdf")
    pdf.save(output_pdf)
    pdf.close()
    print(f"Merged PDF saved as: {output_pdf}")
    os.startfile(output_pdf)
else:
    print("No valid data found. PDF not created.")
