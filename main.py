import os
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill

# Constants for CSV file
CSV_FILE_PATH = 'LogLookupReport_MR.csv'
CSV_HEADER_ROW = 2

# Load the CSV file into a Pandas DataFrame with header starting from row 4
df = pd.read_csv(CSV_FILE_PATH, header=CSV_HEADER_ROW)
print(df.head())

# Parse the 'DateTime' column with the correct format
df['DateTime'] = pd.to_datetime(df['DateTime'], format='%I:%M %p')

# Create a new Excel workbook
output_excel_file = 'pick_counts.xlsx'
book = openpyxl.Workbook()


# Import and perform analysis for different pick types
from putwall_pick import perform_putwall_pick_analysis
perform_putwall_pick_analysis(df, book)

from regular_pick import perform_regular_pick_analysis
perform_regular_pick_analysis(df, book)

from single_pick import perform_single_pick_analysis
perform_single_pick_analysis(df, book)

from replenishment_pick import perform_replenishment_pick_analysis
perform_replenishment_pick_analysis(df, book)

from quick_move import peform_quick_move_analysis
peform_quick_move_analysis(df, book)

from idle_time import perform_idle_time_analysis
perform_idle_time_analysis(df, book)

from hourly_pick_totals import perform_hourly_pick_totals_analysis
perform_hourly_pick_totals_analysis(df, book)

# Save the Excel file
book.save(output_excel_file)

# Check if the file was saved successfully
if os.path.exists(output_excel_file):
    print(f"Excel file '{output_excel_file}' saved successfully.")
else:
    print(f"Failed to save Excel file '{output_excel_file}'.")
