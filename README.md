# OCC Automation Script 

A powerful Selenium-based Python script that automates the validation of customer billing data in a telecom web application. It retrieves detailed charge information and writes results into Excel sheets. Designed to replace time-consuming manual validation tasks, the script increases efficiency and reduces processing time by more than 90%.


## Description

Before automation, the validation of post-production accounts was entirely manual and slow. Each account could include multiple customer accounts (400–500 accounts per task), manual handling was inefficient.

This script automates the login process, customer navigation, OCC (Other Credits and Charges) extraction, unbilled usage scraping, and Excel reporting — all using interactive Selenium automation. It navigates through pages, captures key billing and usage data, and appends results to Excel.

---

## Features

-  Automated login and secure access
-  Fetching OCC (Monthly Fee Adjustments) and unbilled usage
-  Appends results into a structured Excel workbook
-  Summary report generation
-  Saves screenshots on failure
-  Robust error handling with retries

---

## Tools Used

- **Python**
- **Selenium**
- **OpenPyXL**
- **WebDriverWait / ExpectedConditions**
- **Excel Sheet Writing and Formatting**


## How to Run

1. **Install required packages**  
   pip install selenium openpyxl


## Key Logic Summary

The script logs into the billing system using Selenium.
Iterates over account numbers from the Excel sheet.
Navigates the UI to extract:

Monthly fee adjustments
Unbilled usage data

Extracted data is appended to a designated sheet (CX) in the Excel workbook.
Script includes:
Navigation fallback if data is missing
Screenshot saving on failures
Performance marking per account


## Author
Ahmed Essam

