# Financial_Statement_with_Cube_Formula

## From Power Query to Power Pivot to Dynamic Cube Formulas

This project demonstrates how to leverage Power BI’s data model inside Excel to create dynamic, flexible, and fully formatted Financial Statements (Income Statement, Balance Sheet, and Cash Flow Statement).

The process begins with data cleaning and modeling in Power Query and Power Pivot, and culminates in building reusable Cube Formulas that connect directly to the data model — eliminating manual pivot updates and rigid layouts.

## Project Workflow

### Data Cleaning in Power Query
- Source data: Full transactional General Ledger (GL) extracted from the company’s central accounting database.
- Power Query was used to:
    - Remove duplicates and inconsistencies.
    - Format dates and amounts.
    - Establish relationships between key tables — _FactGLTran, DimGLAcct, DimDate, and DimHeaders.
- The cleaned data was loaded into the Data Model.

### Data Modeling in Power Pivot
- Tables were connected via primary and foreign keys.
  
- Calculated columns and key measures were created:
    - [I/S Amount] – Income Statement Measure
    - [B/S Amount] – Balance Sheet Measure
    - [Retained Earnings] – Equity link measure
    - [CF Amount] – Cash Flow Measure
 
- Dimensional tables provided structure for reporting:
    - DimGLAcct – Account classification
    - DimDate – Reporting period
    - DimHeaders – Statement categories (Revenue, Expenses, Assets, Liabilities, Equity
 
### Creating a Pivot Table from the Data Model
- In Excel, a PivotTable was built using the same data model:
    - GLNumber → from FactGLTran
    - GLAcct → from DimGLAcct
    - Amount → [I/S Amount]
 
- This pivot table served as the foundation for Cube Formulas.

### Converting Pivot Table to Cube Formulas
- Select the PivotTable.
- Navigate to: PivotTable Analyze → OLAP Tools → Convert to Formulas
- Excel replaces pivot cells with dynamic Cube Formulas linked to the Data Model.

### Why Cube Formulas?
Unlike traditional PivotTables, Cube Formulas:
- Allow free-form layouts (no pivot restrictions).
- Enable custom formatting of financial reports.
- Dynamically link to model data using named ranges or cell references.

### Key Cube Formulas Used

### Cube Member Formula
Defines the context (dimension or hierarchy) for calculations.

Dynamic Year Cells:

    2021 → =CUBEMEMBER("ThisWorkbookDataModel", "[DimDate].[Year].[All].[" & SelectedReportingYear & "]")
    2020 → =CUBEMEMBER("ThisWorkbookDataModel", "[DimDate].[Year].[All].[" & SelectedReportingYear - 1 & "]")
    2019 → =CUBEMEMBER("ThisWorkbookDataModel", "[DimDate].[Year].[All].[" & SelectedReportingYear - 2 & "]")

### Cube Value Formula
Returns numeric values from the model based on the context.
Income Statement:

    =CUBEVALUE(
    "ThisWorkbookDataModel",
    "[Measures].[I/S Amount]",
    "[DimGLAccts].[GLAcctName].[All].[Revenue]",
    D$8
      )

Here, D$8 refers to the dynamic year cell created using the CUBEMEMBER formula above

To make the formula reusable and dynamic:

    =CUBEVALUE(
    "ThisWorkbookDataModel",
    "[Measures].[I/S Amount]",
    "[DimGLAccts].[GLAcctName].[All].[" & $B10 & "]",
    D$8
    )

Revenue:

      =CUBEVALUE("ThisWorkbookDataModel","[Measures].[I/S Amount]","[DimGLAccts].[GLAcctName].[All].[Revenue]",D$8)


Cost of Sales: Same as above referencing $B10

Operating Expenses:

      =CUBEVALUE("ThisWorkbookDataModel","[Measures].[I/S Amount]","[DimGLAccts].[Subcategory].[All].[" & $B22 & "]","[DimHeaders].[Category].[All].[Operating Expenses]",D$8)

Interest Expense: Similar formula with “Interest Expenses” category

Income Tax Expense: “Income Tax Expenses” category

![IS](https://github.com/adetonayusuf/Financial_Statement_with_Cube_Formula/blob/main/Income%20Statement%20-%20Cube.png)


### Balance Sheet

Current Assets:

      =CUBEVALUE("ThisWorkbookDataModel","[Measures].[B/S Amount]","[DimGLAccts].[Subcategory].[All].[" & $B44 & "]","[DimHeaders].[Category].[All].[Current Assets]",D$8)

Non-Current Assets: 

    =CUBEVALUE("ThisWorkbookDataModel","[Measures].[B/S Amount]","[DimGLAccts].[Subcategory].[All].[" & $B44 & "]","[DimHeaders].[Category].[All].[Non-Current Assets]",D$8)

Current Liabilities:

    ...["Current Liabilities"]...

Non-Current Liabilities:

    ...["Non Current Liabilities"]...

Share Capital:

    ...["Equity"]...

Retained Earnings

    =CUBEVALUE("ThisWorkbookDataModel","[Measures].[Retained Earnings]",D$8)

![BS](https://github.com/adetonayusuf/Financial_Statement_with_Cube_Formula/blob/main/BS%20-%20Cube.png)


### Cash Flow Statement

Opening Balances 

      =CUBESET("ThisWorkbookDataModel","[FactGLTran].[GLTranDescription].[ALL].[Opening Balance]","Opening Balance")

  Operating Activities

      Changes in working capital = difference between opening and closing balances of Current Assets & Liabilities.

  Investing Activities

      =-CUBEVALUE("ThisWorkbookDataModel","[Measures].[CF Amount]","[FactGLTran].[GLTranDescription].[All].[Purchase PPE]","[DimGLAccts].[GLAcctName].[All].[Property, Plant & Equipment]",D41)

  Financing Activities

      Derived from movement in equity accounts and retained earnings.

![CF](https://github.com/adetonayusuf/Financial_Statement_with_Cube_Formula/blob/main/cf%20-%20cube.png)

  ### Key Takeaways
  - Cube Formulas bridge Power BI and Excel, offering dynamic control over financial report layouts.
  - Reusable formulas reduce manual intervention while keeping design flexibility.
  - Dynamic year selection allows period comparison across multiple years instantly.
  - Enables audit-friendly visibility since every figure traces back to the Power BI data model.

### Tools & Technologies

  - Excel (Power Query, Power Pivot, OLAP Tools, Cube Formulas)
  - SQL Server / Accounting Database
  - Power BI (Data Model Integration)
  - DAX Measures

### Outcome
The final workbook delivers:
- Dynamic Income Statement, Balance Sheet, and Cash Flow Statement.
- Instant year-to-year analysis without rebuilding pivots.
- A robust bridge between Power BI analytics and Excel reporting flexibility.

