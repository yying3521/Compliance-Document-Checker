import os
import pandas as pd
from openpyxl import load_workbook

# Function to check files in each folder
def check_folders(base_folder):
    output = []  # List to store results for each folder

    # Loop through all subfolders
    for folder_name in os.listdir(base_folder):
        folder_path = os.path.join(base_folder, folder_name)
        if os.path.isdir(folder_path):  # Check if it's a folder
            has_excel = False
            has_pricing = False
            invalid_files = []
            corrupted_files = []

            # Check all files in the folder
            for file_name in os.listdir(folder_path):
                file_path = os.path.join(folder_path, file_name)

                # Check if the file is an Excel file
                if file_name.endswith((".xlsx", ".xls")):
                    has_excel = True
                    if "Pricing" in file_name:
                        try:
                            load_workbook(file_path)  # Try to open the file
                            has_pricing = True
                        except Exception:
                            corrupted_files.append(file_name)  # File is corrupted
                elif "Pricing" in file_name:
                    invalid_files.append(file_name)  # Non-Excel files with 'Pricing'

            # Determine the folder status
            if not has_excel:
                status = "No Excel files found"
            elif not has_pricing:
                status = "No Excel files containing 'Pricing'"
            elif corrupted_files:
                status = f"Corrupted Excel files: {', '.join(corrupted_files)}"
            else:
                status = "Valid Excel files found"

            # Append folder results
            output.append({
                "Folder": folder_name,
                "Status": status,
                "Invalid Files": ", ".join(invalid_files) if invalid_files else "None"
            })

    # Create a DataFrame and save it as a CSV file
    output_df = pd.DataFrame(output)
    output_csv = os.path.join(base_folder, "output_report.csv")
    output_df.to_csv(output_csv, index=False)
    print(f"Report generated: {output_csv}")

# Set the path to the main folder
base_folder = os.path.expanduser("~/Desktop/Proposal-Sample-Set")  # Automatically get the user's desktop path
check_folders(base_folder)


