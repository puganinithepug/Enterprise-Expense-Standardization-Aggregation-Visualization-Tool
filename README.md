# Enterprise 10-Q Reporting 

## Standardization Aggregatoin and KPIs for Expense Reporting - VBA Macro

This VBA automation data pipeline is designed to streamline enterprise expense tracking and quarterly reporting by consolidating data across multiple worksheets into a single, standardized quarterly report.

_What traditionally requires manual copying, formatting, header management, and recalculation across multiple sheets is reduced to an straightforward click-through execution._

## By automating repetitive, error-prone tasks, this solution effectively decreases tedious effort and introduces consistency, scalability, and audit-readiness into financial reporting workflows

**_Run VBA Macro LaunchApp - Starts the Main User Form (Quarterly Expenses)_**

**Outcomes:**

- "Raw Data Quarter X Expenses i" sheets: **Data Simulation/Generation**
- "Quarter X Expenses" sheet: **Data Standardization/Aggregation**
- "Pivot Table for Quarter X" pivot table sheet: **Data Transformation**
- _Where  X = quarter, i = index_

1. Generates a number (user's choice) of raw data sheets, prepopulated with randomly generated data (Populate Sample Data).
2. Standardizes raw data data for a specific quarter (selected by user in a dropdown menu) - the dropdown menu cannot be used until data is populated.
3. Aggregates data (in the same event as standardizaton) into a "Quarter X Expenses" sheet, where X - 1, 2, 3, 4 based on user selection.
4. Simultaneously with the aggregation/standardization of raw data sheets and the generation of the "Quarter X Expenses" sheet, a pivot table is built (based on the "Quarter X Expenses" sheet) inside of the "Pivot Table for Quarter X". The pivot table shows the "Sum of Total" per Category per Division. In otherwords, there is a total for every Category in a Division, a total for every Division, a grand total for that Quarter. It is a data transformation add-on, to improve structure and readabilty of the "Quarter X Expenses"
- If the user wants to generate more data, say for another quarter, the user must simply click the Populate Sample data button again. This generates more Raw data sheets (however-many user requests - following the workflow from step 1, the previously generated sheets are not overwritten).
- If in following generations of new data, the quarter selected already has correspondng data generated, then the user will be given the option to overwrite data or keep existing - in that case workflow is exited.
4. There is a button to refresh workbook, so all sheets with data are deleted and there is only 1 empty sheet left (Sheet1 - to mimic default workbook set up).
5. There is a Close button that will exit out of all user forms, exit the workflow. The sheets generated are preserved.
6. The Quick Analysis button launches the Quick Analysis user form. The generated sheets are preserved.

**_Pressing the Quick Analysis Button Starts the Lookup User Form (Quick Analysis)_**

**Outcomes:**

- Sum or average or standard deviation per specific Division & Category selection, across specific quarters: **Data Analysis, KPIs**

1. The main button at the top of user form (Run Lookup) is disabled at the beginning. This is the button that renders the results. Prior to using this button the user must make selections on the form.
2. The user must select both a Division (east, west, north, south) and a Category (any of the possible categories available as possible expenses, not guaranteed to be in the generated data set) from the dropdowns.
3. The user must also check off at least one quarter: Check off which quarters to include in the analysis: Q1, Q2, Q3, Q4. If a quarter that has no corresponding data sheets in the workbook the user recieves a message box pop up with "Yes" and "No" options. If the user wants corresponding raw data sheets and "Quarter X Expenses" sheet generated, they must select "Yes", which automatically renders the raw data sheets, the "Quarter X Expenses" sheet and the corresponding pivot table sheet (basically generates outcomes of the Main user form for additional workflow flexibility).
4. User must select one KPI - Sum or Average or Standard Deviation for Expenses matched for selected Division/Category lookup key, per across selected quarter(s). KPI selection is mutually exclusive - either calculate Sum, or Average, or Standard Deviation.n This is enforced by deselection of the initially selected KPI, if the user attempts to select a second KPI.
5. After the user currectly selects all preferences from the user form, the Run Lookup button is enabled, it can be pressed to generate a numeric result in the text box below the Result label.
- _Note:_ The text box will show "No matching data" when the generated data for that quarter (or across multiple quarters) does not contain a value for the Division/Category composite key, so it's not possible to generate a sum/average/standard deviation value. I.e there were no Expenses classified to the selected Category for that Divison, in that quarter (quarters).
6. After pressing the Run Lookup button and a result appears in the text box, the Run Lookup button is temporarily disabled, the dropdowns and check boxes will be temporarily disabled as well. The user must press the Clear Selection button to clear the previous selections on the user form. After pressing the Clear Selection button, the dropdowns. check boxes are enabled again, the Run lookup remains disabled until minimum selections are made on the user form (as discussed in previous steps).
7. The Back To Data button takes the user back to the previous user form (Main user form).
8. There is a Close button that will exit out of all user forms, exit the workflow. The sheets generated are preserved.

_Note: the only way to clear the workbook to original state is to use the Refresh Workbook button on the Main user form._

_Development note: VLOOKUP -> XLOOKUP -> INDEX/MATCHING:_

- The initial plan was to implement VLOOKUP for retrieval of the Total based on a combined key composed from Division and Category, searching only in selected quarters.
- This approach was unideal because it would requiring copying data - ensuring that the value column to be searched is on the righthand side of the lookup key column.
- Since the composite key is generated after the initial data set and is placed (as a hidden and protected column) on the righthand side of the Total column (the value column), the plan to use VLOOKUP was strategically adjusted. The best replacement for VLOOKUP would be XLOOKUP which is not contrained to column positions. Since the version of Excel used is 2016, it does not support XLOOKUP.
- The workaround was using INDEX/MATCH instead, it is not constrained to specific column positions for lookup key and value.
- When matches are found for the composite key, results are collected and processed based on what KPIs (sum, average or standard deviation) the user wants to calculate.
- The result appears as output.

## 10-Q Sales, Revenue and Expenses Visualizations for Business Insights - Power BI & Microsoft Fabric

The second part of this Enterprise 10-Q ETL is visualization of enterprise sales, revenue and costs using Power BI. Made as a follow up to the VBA Macro tool

The Power BI report is designed for easy navigation for readers. The report also leverages dynamic RLS (role level security) to ensure controlled user access priviledges. The primary report page is the _Page Navigation_ page. This page is directly integrated with RLS, ensuring controlled access. The page has a selection panel for various pages in the report containing graphical analysis of the enterprise finances. Once a panel selection is made, the button click redirects the user to the selected report page.
- RLS related data is contained in the **Security Table** and the **PLS** datasets.
- The **Security Table** contains a column with user identities, a column for user emails, and a column with a state name to indicate which region's data the user has access to across the report - relative to their role in the enterprise. This is specifically demonstrating how the data access can be partitioned for managers of different divisions within the enterprise.
- The **PLS** is a table containing a column with emails and a column called Page Access with the names of pages which the user of the email has access to within the report.
- In both cases dynamic RLS works by detecting ehich user is signed in, to adjust the visbility of report data.

_Data Transformation_
- DAX Query was essential for exploring raw imported dataset, detecting problematic areas, as well as for creating measures based on the data for deriving KPIs.
- Raw CSV and excel spreadsheets are imported and transformed with Power Query editor prior to any visualizations or analysis.
- Maintaining the data model was an importeant component for organizing relationships between related datasets. 

_Data Tables of the Report_
- The **sales_total_2** table has the order_id  column as its primary key, with other columns containing the count of sales, the price, the revenue, the stock, the store id. the product_id, order_date, promo_type, promo_bin, promo_discount and Discount Price (a calculated column with DAX). This is the fact table.
- The **producthierarchy** table contains product specific details. The product_id keys relate the producthierarchy table to the sales_total_2 table. This table is a dimension table for product-related information. This table contains columns: category, sub-category, product (brand), type, length, width, depth and Volume - a calculated column derived from length/depth/width. 
- The **store_cities** table contains information about the stores and their geographical location. This table is a dimension table for store information. The store_id keys relate the store_cities table to the sales_total_2 table. This table contains columns: city, city_id, cost, cost type, latitude, longitude, state, state_id, store size in m^3, state abbreviation, and store_type_id.
- The **DateTable** is a dimension table for order dates. This table relates to the fact table through the order_date column. This table contains columns such as: weekday, date, date as integer, quarter, month number, month short, year month, etc. It was useful for formatting different visuals.
- The **Measure table** contained differet derived statistics and KPIs which were essenntial for visualizations:

      total revenue, Base Pkg Dimension, Cost of Order, Cost of Order Filter California, AVG Product Price 7%, AVG layout dim, Logistic Category Costs, Logistic Category Costs with Volume, All Costs, % Cost, % Cost state only, Cost Previous, Change in Cost, YTD Cost, YTD Revenue, Monday sales shares, All Revenue, % Revenue by Category, % Revenue Beauty & Hygiene, and YTD Revenue till August 2018...

_Data Visualization_
- The report contains several pages which demonstrate different methods of analyzing revenue and costs associated with enterprise operations from 2017 to 2019
- **Unpivot** page contains unpivoted cost data, where costs are displayed in a stacked column chart, columns are organized by store_id, and the portion of cost is color coded by cost type within each column. There is also a bar chart for sum of revenue per state. This page also contains 2 slicers. One for filtering by state and within each state by city for stores in the store_id to cost bar chart; the second a slider for filtering by interval of price range of products sold, to be included in the revenue aggregation.
- **Details** page contains date based KPIs with formatted dates. There is also a bar graph which displays revenue by weekdays. The page also displays the total revenue and the total sales. 


