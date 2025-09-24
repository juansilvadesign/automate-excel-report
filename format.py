# Adds Visual Formatting Only
# Purpose: Takes an existing report and adds visual styling

# Reads from report.xlsx
# Adds title "Sales Report" and "Month" labels
# Applies custom fonts (Rajdhani, Inter) and colors
# Saves as report_month.xlsx

from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.styles.colors import Color
import os

wb = load_workbook('report/report.xlsx')
sheet = wb.active  # Use active sheet instead of hardcoded 'Report'

sheet['A1'] = 'Sales Report'
sheet['A2'] = 'Month'
sheet['A1'].font = Font('Arial', bold=True, size=20, color=Color(rgb='85B50B'))
sheet['A2'].font = Font('Arial', size=10)

# Create report directory if it doesn't exist
os.makedirs('report', exist_ok=True)
wb.save('report/report_month.xlsx')