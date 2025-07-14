import glob
import os
from datetime import datetime

import pandas as pd

# Define paths
disabled_path = os.path.join(os.path.expanduser("~"), "Downloads/EPDR_Reports/Disabled")
offline_path = os.path.join(os.path.expanduser("~"), "Downloads/EPDR_Reports/Offline")


def excel_generator(folder_path, report_name):
    csv_files = glob.glob(os.path.join(folder_path, "*.csv"))

    combined_df = pd.DataFrame()

    for csv_file in csv_files:
        try:
            df = pd.read_csv(csv_file, delimiter="\t", encoding="utf-16")

            df_filtered = df.iloc[:, [2, 9]].copy()
            df_filtered.columns = ["Device", "Group"]
            df_filtered["Revised by"] = ""  # Empty column for manual entry
            df_filtered["Case number"] = ""

            # Add identifier from filename (without .csv)
            csv_name = os.path.splitext(os.path.basename(csv_file))[0]
            df_filtered.insert(0, "Account", csv_name)

            # Append to combined dataframe
            combined_df = pd.concat([combined_df, df_filtered], ignore_index=True)

        except Exception as e:
            print(f"Failed to process {csv_file}: {e}")

    if not combined_df.empty:
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        output_excel = os.path.join(folder_path, f"{report_name}_{timestamp}.xlsx")
        combined_df.to_excel(output_excel, index=False)
        print(f"Excel report saved as: {output_excel}")
        os.startfile(output_excel)
    else:
        print("No valid data found. Excel file not created.")


# Generate both Excel reports
excel_generator(disabled_path, "EPDR_Disabled_Protection_Report")
excel_generator(offline_path, "EPDR_Offline_Devices_Report")
