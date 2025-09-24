from openpyxl import Workbook

# Create a new workbook and worksheet
wb = Workbook()
ws = wb.active
ws.title = "Sales Data"

# Sample sales data - Products in Column A, Sales data in other columns
data = [
    ["Product", "Q1 Sales", "Q2 Sales", "Q3 Sales", "Q4 Sales", "Total Units"],
    ["Laptops", 1250, 1380, 1150, 1420, 5200],
    ["Smartphones", 2100, 2250, 1950, 2400, 8700],
    ["Tablets", 850, 920, 780, 1100, 3650],
    ["Headphones", 1450, 1200, 1350, 1600, 5600],
    ["Keyboards", 680, 750, 620, 850, 2900],
    ["Monitors", 920, 1100, 980, 1200, 4200],
    ["Webcams", 560, 620, 480, 750, 2410],
    ["Speakers", 780, 850, 720, 950, 3300],
    ["Mouse", 950, 1050, 890, 1200, 4090],
    ["Printers", 420, 380, 450, 520, 1770]
]

# Write data to worksheet
for row in data:
    ws.append(row)

# Save the workbook
wb.save("raw_data.xlsx")
print("✅ Created raw_data.xlsx with sample sales data")
print("\nData structure:")
print("- Column A: Product names")
print("- Columns B-F: Quarterly sales data")
print("- 10 products with realistic sales numbers")
print("- Headers in Row 1")
print("- Ready for automation testing!")