# Contract and Invoice Data Comparison Tool

## Description
This Python script provides a solution for comparing contract and invoice data from an Excel file. It calculates fuzzy matching scores between descriptions and part numbers to identify discrepancies between these entries. The discrepancies found are saved into a new sheet within the same Excel file, allowing for easy review and further processing.

## Features
- **Data Loading**: Load contract and invoice data from specified sheets within an Excel file.
- **Fuzzy Matching**: Calculate similarity scores using `difflib.SequenceMatcher` to compare descriptions and part numbers between contracts and invoices.
- **Discrepancy Detection**: Identify and record discrepancies based on fuzzy scores and quantity mismatches.
- **Excel Integration**: Output the discrepancies to a new sheet in the same Excel file.

## Installation

### Prerequisites
Ensure you have Python installed on your machine. Python 3.6 or later is recommended. You will also need `pandas` and `openpyxl` libraries for handling Excel file operations.

### Setup
1. Clone this repository to your local machine using:

  ```bash
  git clone https://github.com/your-github-username/contract-invoice-comparison.git
```
2. Install the required Python libraries:

  ```bash
  pip install pandas openpyxl
```

## Usage
### To use this script:

Place your Excel file in the same directory as the script, or modify the FILE_PATH variable in the script to point to the location of your Excel file.
Run the script using:

```bash
python contract_invoice_comparison.py
```

Check the same Excel file for a new sheet named "Discrepancies" which will contain all identified discrepancies.
## Author
**Carlos Germosen Polanco**

Date: June 19, 2024
