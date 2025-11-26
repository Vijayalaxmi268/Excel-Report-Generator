# 📊 Excel Report Generator (Python + Tkinter + Excel)

## 🎯 Objective

This project is a desktop tool that **automatically generates Excel reports from CSV files**.  
Using a simple Tkinter GUI, the user can:

- Select a CSV file
- Automatically create **pivot tables** and **summary statistics**
- Generate a **bar chart** using matplotlib
- Export everything into a **styled Excel workbook** (`.xlsx`) using openpyxl

It is a mini end-to-end data analysis + reporting automation project.

---

## 🧰 Tools & Technologies

- **Python 3**
- **pandas** → Load & analyze CSV data
- **matplotlib** → Create charts
- **openpyxl** → Create & style Excel reports, embed chart images
- **tkinter** → GUI dialogs for file open/save

---

## 📁 Project Structure

Example structure:

```bash
excel_report_generator/
│
├── report_generator.py        # Main Python script (GUI + logic)
├── sample_data.csv      # Large dataset with 100+ rows
├── output/                    # Folder for exported Excel reports (optional)
├── screenshots/               # Screenshots of GUI + Excel output (for submission)
└── README.md                  # Project documentation

The script is designed to work with a CSV that has these columns:
 Date, Region, Product, Category, Sales, Profit, Quantity
 Date → e.g. 2025-01-01
 Region → North, South, East, West, etc.
 Product → Apples, Oranges, Coffee, Carrots, etc.
 Category → Fruit, Vegetable, Beverage, etc.
 Sales → Numeric
 Profit → Numeric
 Quantity → Numeric

# You can use the provided sample_data.csv or create your own CSV with the same columns.

## ⚙️ Installation

1 Make sure Python 3 is installed.
2 Install required libraries:
3 pip install pandas matplotlib openpyxl
4 tkinter usually comes pre-installed with Python on Windows.
5 If not, install it via your OS package manager.

## 🚀 How to Run the Application

1 Open a terminal / command prompt in the project folder.

2 Run:
    python report_generator.py

3 The Tkinter file dialog will open:
    Select your CSV file (e.g. sample_data_large.csv).

4 The script will:
    Load the CSV with pandas
    Create pivot tables and summary stats
    Generate a bar chart image
    Build a styled Excel workbook

5 A Save As dialog will appear:
Choose a location and name for the Excel report, e.g.:
   output/sales_report.xlsx

6 On success, you’ll see a popup:

✅ "Report saved to: …"