## Data Automation Developer: Excel Workflow Streamlining

This repository contains the Python automation scripts designed to streamline payroll and HR data processing.

The goal of this solution is to reliably extract data from master files, transform it by grouping and cleaning, and load it into specific per-employee Excel reports while preserving existing formulas and formatting.

---

### Features

* **ETL Workflow:** Implements a complete Extract, Transform, Load (ETL) pipeline using Python. 
* **Robust Excel Handling:** Utilizes `pandas` and `openpyxl` to append new data to existing Excel files without corrupting sheet structure, formulas, or formatting.
* **Modular Design:** Organized into clear functions (`load_and_clean_master_data`, `append_to_employee_file`) for maintainability and readability.
* **Automated Grouping:** Automatically groups the master dataset by `Employee ID` to facilitate targeted data loading.
* **Comprehensive Logging:** Tracks all processing steps, warnings (e.g., missing employee files), and errors to a dedicated log file, ensuring auditability.
* **Path Management:** Uses `pathlib` for robust, cross-platform file and directory management.

---

### Prerequisites

To run this automation script, you need Python and the necessary libraries installed.

* Python 3.8+

#### Dependencies

Install the required libraries using `pip`:


```
pip install pandas openpyxl 
```
## Project Structure 

The project relies on a specific folder structure for inputs and outputs.
```

data_automation_project/
├── master_data/                 # INPUT: Source Excel files (e.g., payroll_update_YYYYMMDD.xlsx)
├── employee_files/              # INPUT/OUTPUT: Existing per-employee workbooks (EMP_XXXX_Report.xlsx)
├── automation_script.py         # The core ETL script
└── automation_log_YYYYMMDD.log  # OUTPUT: Log file created upon execution
```
