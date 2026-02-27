# Manufacturing Automation Suite
A Python-based document automation system that eliminates manual data entry 
and error-proofs critical manufacturing documents. Driven by a central 
dataframe, the suite generates formatted Excel workbooks for process logging 
and work order management.

---

## The Problem This Solves
In manufacturing environments, process logs and work orders are typically 
created manually — copying specs from reference documents, filling in 
parameters by hand, and relying on multiple sign-off tiers to catch errors. 
This is time-consuming and error-prone.

This suite automates that entire workflow. As long as the source dataframe 
is accurate, every generated document is guaranteed to have the correct 
specifications for the selected product and line speed — with zero manual 
lookup required.

---

## Programs

### `data.py`
The data backbone of the suite. Contains two pandas DataFrames:
- **Production DataFrame** — stores all machine parameters, raw material 
  specs, quality tolerances, process times, and BOM data by product and 
  line speed
- **Label DataFrame** — stores customer, dimensions, weight, and label 
  requirements per product

### `log_gen.py`
Generates a blank, fully formatted process log template in Excel with:
- Info block (product, operator, shift, date)
- Raw material checks section with Pass/Fail columns
- Machine parameter checks section with Pass/Fail columns
- Quality summary with automatic PASS/FAIL counts
- Supervisor sign-off block

### `log_main.py`
Populates the process log template with product-specific data. User selects 
product and line speed from a menu, enters shift details, and the program 
generates a ready-to-use process log pre-loaded with all correct 
specifications and tolerances.

### `work_order.py`
Generates a three-sheet Excel workbook from a single user session:
- **Work Order** — title block, production summary with shift calculations, 
  bill of materials, and post-processing instructions
- **Setup Sheet** — all machine parameters with set points and 
  min/max tolerances
- **Label Info** — customer details, finished dimensions, and label 
  requirements for all packaging stages

---

## Tech Stack
- Python 3.x
- pandas
- openpyxl

---

## Setup
```bash
# Clone the repository
git clone https://github.com/Brett-Beery/manufacturing-automation.git
cd manufacturing-automation

# Create and activate virtual environment
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# Install dependencies
pip install pandas openpyxl
```

---

## Usage
```bash
# Generate a process log
python log_main.py

# Generate a work order
python work_order.py
```

---

## Project Background
This project is a sanitized demonstration of a real automation system 
built for a production manufacturing environment. The original system 
eliminated an estimated 20-30 hours of weekly clerical work, removed 
3 tiers of manual sign-offs, and error-proofed critical production 
documents across multiple machine centers.

Dummy products and placeholder data have been substituted for all 
proprietary information.

---

## Author
Brett Beery — Python developer specializing in automation and data 
manipulation. [GitHub](https://github.com/Brett-Beery) | 
[LinkedIn](https://www.linkedin.com/in/brett-beery/)