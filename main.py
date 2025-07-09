import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename

import fitz  # PyMuPDF
import pandas as pd


def main():
    # "C:/Users/nj250196/Downloads"
    downloads_folder_path = os.path.join(os.path.expanduser("~"), "Downloads")

    # Open file picker dialog (Downloads folder as initial directory)
    Tk().withdraw()  # Hide the root window
    csv_file = askopenfilename(
        title="Select a CSV file",
        filetypes=[("CSV files", "*.csv")],
        initialdir=downloads_folder_path,
    )

    # Load the CSV file
    df = pd.read_csv(csv_file, delimiter="\t", encoding="utf-16")

    # Select only the 3rd and 10th columns
    df_filtered = df.iloc[:, [2, 9]]

    # PDF generation
    pdf = fitz.open()
    page_width, page_height = fitz.paper_size("a4")
    margin = 50
    line_height = 18
    column_spacing = 150
    max_lines_per_page = int((page_height - 2 * margin) / line_height)

    def add_page_with_rows(rows):
        page = pdf.new_page(width=page_width, height=page_height)
        for i, (col1, col2) in enumerate(rows):
            y = margin + i * line_height
            page.insert_text((margin, y), str(col1), fontsize=10, fontname="helv")
            page.insert_text(
                (margin + column_spacing, y), str(col2), fontsize=10, fontname="helv"
            )

    rows = [tuple(df_filtered.columns)]
    for _, row in df_filtered.iterrows():
        rows.append((row.iloc[0], row.iloc[1]))

    for i in range(0, len(rows), max_lines_per_page):
        add_page_with_rows(rows[i : i + max_lines_per_page])

    output_pdf = csv_file.replace(".csv", "_filtered.pdf")
    pdf.save(output_pdf)
    pdf.close()

    print(f"PDF saved to: {output_pdf}")

    # Open PDF after conversion
    os.startfile(output_pdf)


if __name__ == "__main__":
    main()
