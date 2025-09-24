# 📊 Excel Report Automation

An automated Excel report generation system built with Python and openpyxl. This project transforms raw sales data into professional reports with charts, formulas, and formatting.

## 🎯 Project Purpose

I created this project to **train my Python skills** using the **openpyxl library**. It demonstrates practical automation techniques for Excel file manipulation, including:
- Reading and writing Excel files
- Creating dynamic charts and visualizations
- Implementing formulas programmatically
- Applying professional formatting and styling

## ✨ Features

- **📈 Automated Bar Charts** - Visual representation of sales data by product line
- **🧮 Dynamic Formulas** - Automatic SUM calculations for all data columns
- **🎨 Professional Formatting** - Custom fonts, colors, and styling
- **📁 Organized Output** - All reports saved to dedicated `/report` folder
- **🔧 Modular Design** - Separate scripts for different automation tasks
- **💻 Executable Ready** - Can be converted to standalone .exe file

## 🚀 Quick Start

### Prerequisites

- Python 3.7+
- pip package manager

### Installation

1. Clone the repository:
```bash
git clone https://github.com/juansilvadesign/automate-excel-report.git
cd automate-excel-report
```

2. Create and activate virtual environment:
```bash
pip install virtualenv
virtualenv .env
.env\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Run the sample data generator:
```bash
python create_sample_data.py
```

4. Run the main automation:
```bash
python py-to-exe.py
```

## 📁 Project Structure

```
📁 Automate Excel Report/
├── 📄 raw_data.xlsx           # Input data (sample included)
├── 📄 py-to-exe.py            # Main automation script
├── 📄 formula.py              # Formulas only
├── 📄 format.py               # Formatting only  
├── 📄 formula+format.py       # Complete automation
├── 📄 create_sample_data.py   # Sample data generator
├── 📄 requirements.txt        # Dependencies
└── 📁 report/                 # Output folder
    ├── 📄 report_january.xlsx
    ├── 📄 report.xlsx
    └── 📄 report_month.xlsx
```

## 📊 Input Data Format

Your Excel file must follow this structure:

| Product | Q1 Sales | Q2 Sales | Q3 Sales | Q4 Sales |
|---------|----------|----------|----------|----------|
| Laptops | 1250 | 1380 | 1150 | 1420 |
| Smartphones | 2100 | 2250 | 1950 | 2400 |
| Tablets | 850 | 920 | 780 | 1100 |

**Requirements:**
- Column A: Product names
- Columns B+: Numerical sales data
- Row 1: Headers
- No empty rows/columns

## 🛠️ Usage Options

### 1. Complete Automation (Recommended)
```bash
python py-to-exe.py
```
- Interactive month input
- Creates bar charts, formulas, and applies formatting
- Output: `report/report_{month}.xlsx`

### 2. Individual Components

**Add formulas only:**
```bash
python formula.py
```
Output: `report/report.xlsx`

**Add formatting only:**
```bash
python format.py
```
Output: `report/report_month.xlsx`

**Complete automation (hardcoded):**
```bash
python formula+format.py
```
Output: `report/report_january.xlsx`

## 📈 Output Features

Generated reports include:

- **📊 Bar Chart**: Visual sales comparison by product
- **💰 SUM Formulas**: Automatic totals with currency formatting
- **🎨 Professional Styling**: 
  - Custom title formatting
  - Month labels
  - Color-coded elements
- **📅 Month-based Naming**: Dynamic file naming

## 🔧 Creating Executable

Convert to standalone .exe file:

```bash
pip install pyinstaller
pyinstaller --onefile py-to-exe.py
```

## 📚 Learning Outcomes

Through this project, I gained hands-on experience with:

- **openpyxl fundamentals**: Reading, writing, and manipulating Excel files
- **Chart creation**: Using `BarChart` and `Reference` objects
- **Formula implementation**: Dynamic formula generation with `get_column_letter`
- **Styling and formatting**: Working with `Font`, `Color`, and cell formatting
- **File path handling**: Cross-platform compatibility and executable deployment
- **Python best practices**: Modular code structure and error handling

## 🤝 Contributing

Feel free to fork this project and submit pull requests for improvements!

## 📄 License

This project is open source and available under the [MIT License](LICENSE).

## 🙏 Acknowledgments

- Built with [openpyxl](https://openpyxl.readthedocs.io/) - The amazing Python library for Excel files
- Created as a learning exercise to master Python automation techniques

---

**Happy Automating!** 🚀

*This project demonstrates practical Python skills for Excel automation using openpyxl.*