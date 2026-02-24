# Enterprise Expense Standardization, Aggregation & Visualization Tool

This VBA automation tool is designed to streamline enterprise expense tracking and quarterly reporting by consolidating data across multiple worksheets into a single, standardized quarterly report.

What traditionally requires manual copying, formatting, header management, and recalculation across multiple sheets is reduced to a one-click execution.

By automating repetitive, error-prone tasks, this solution effectively decreases tedious effort and introduces consistency, scalability, and audit-readiness into financial reporting workflows

- User-driven quarter selection (Q1–Q4)
- Automatic worksheet creation for quarterly expense reports
- Dynamic month headers aligned to the selected quarter
- Automated total calculations with safeguards against duplication
- Header detection logic to prevent stacking on reruns
- Cross-worksheet data consolidation
- Rerun-safe design suitable for iterative reporting cycles

**Accessible Tool Navigation**
- The script allows the user to easily navigate through different sheets of the workbook via user form that contains a drop down of all sheets in the workbook.

**Using the Tool**
- Additionally the user form has a button that allows the creation of new sheets in he workbook and their renaming.
Likewise, the script allows the user to import files into excel, by prompting the user to (multi) select files to import.
- The user form also has a button allowing the user to run the yearly report on data imported into the workbook given the same core staructure.

**Automated Data Standardization**
- The script automates the addition of appropriate headers (given similarly structured data) and formatting of the financial data as currency.

**Data Aggregation**
- Running the yearly report will compile all data from the sheets in the workbook and paste it into the yearly report sheet.
The aggregated data sheet produced is also formatted.

