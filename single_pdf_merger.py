import glob
import os

import fitz  # PyMuPDF
import pandas as pd

disabled_path = os.path.join(os.path.expanduser("~"), "Downloads/EPDR_Reports/Disabled")
offline_path = os.path.join(os.path.expanduser("~"), "Downloads/EPDR_Reports/Offline")


def pdf_generator(folder_path):
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
        nonlocal current_page, current_line, pages_added
        if current_page is None or current_line >= max_lines_per_page:
            current_page = pdf.new_page(width=page_width, height=page_height)
            current_line = 0
            pages_added += 1

    def write_row(col1, col2, align="left", fontname="helv", fontsize=10):
        nonlocal current_line
        ensure_page()
        y = margin + current_line * line_height

        if align == "center":
            text = str(col1)
            try:
                text_width = fitz.get_text_length(
                    text, fontsize=fontsize, fontname=fontname
                )
            except:
                text_width = 0  # Safe fallback
            x = (page_width - text_width) / 2
            current_page.insert_text((x, y), text, fontsize=fontsize, fontname=fontname)
        else:
            current_page.insert_text(
                (margin, y), str(col1), fontsize=fontsize, fontname=fontname
            )
            current_page.insert_text(
                (margin + column_spacing, y),
                str(col2),
                fontsize=fontsize,
                fontname=fontname,
            )

        current_line += 1

    for csv_file in csv_files:
        try:
            df = pd.read_csv(csv_file, delimiter="\t", encoding="utf-16", skiprows=1)
            if df.shape[1] < 10:
                print(f"Skipping {csv_file}: not enough columns.")
                continue

            df_filtered = df.iloc[:, [2, 9]]
            csv_name = os.path.basename(csv_file)

            write_row("", "")  # Space above title
            write_row(
                f"--- {csv_name} ---", "", align="center", fontname="helv", fontsize=12
            )
            write_row("", "")  # Space below title
            write_row(*df_filtered.columns)  # Write column headers

            for _, row in df_filtered.iterrows():
                write_row(row.iloc[0], row.iloc[1])

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


# Generate both PDFs in their respective folders
pdf_generator(disabled_path)
pdf_generator(offline_path)
