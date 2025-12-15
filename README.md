# PR.2 Analyzer - Excel Data Analysis Project

## Project Overview
This project demonstrates comprehensive Excel data analysis skills including conditional formatting, What-If analysis, linear regression, descriptive statistics, and data visualization through dashboards. The focus is on analyzing sales data to derive actionable business insights.

## Topics Covered

### 1. **Conditional Formatting**
- Visual highlighting of key data points
- Top 10 customer identification
- Performance-based color scales

### 2. **WHAT IF Analysis**
- Scenario planning for discount strategies
- Impact assessment on total profit
- Sensitivity analysis

### 3. **Linear Regression (Data Analysis Toolpak)**
- Correlation between Sales and Profit
- Predictive modeling
- Statistical significance testing

### 4. **Data Analysis Features**
- Descriptive Statistics (Mean, Median, Mode, SD)
- Regression analysis
- Histogram generation

### 5. **Data Storytelling**
- Converting raw data into meaningful insights
- Dashboard creation
- Visual narrative building

### 6. **Custom Formatting**
- Symbols and arrows for trend indication
- Up/down arrows for growth metrics
- Enhanced data readability

### 7. **Timestamp Creation**
- Dynamic date-time tracking using NOW()
- Data audit trails

### 8. **High-Value Customer Identification**
- Advanced formula combinations: GROUPBY(), INDEX(), MATCH()
- Customer segmentation
- Revenue concentration analysis

### 9. **Advanced Descriptive Analytics**
- Statistical measures
- Distribution analysis
- Outlier detection

### 10. **Pivot Tables with Dashboard**
- Multi-dimensional data aggregation
- Interactive reporting
- Regional and product performance analysis

### 11. **Data Visualization**
- Bar charts for comparative analysis
- Line charts for trend analysis
- Pie charts for composition analysis

---

## Main Tasks

### Task 1: Conditional Formatting
**Objective:** Highlight the top 10 customers based on total purchase value.

**Steps:**
1. Calculate total purchase per customer
2. Apply conditional formatting rule (Top 10 items)
3. Use color scale or icon sets for visual emphasis

### Task 2: WHAT IF Analysis
**Objective:** Analyze the impact of changing discount rates on total profit.

**Steps:**
1. Create discount scenarios (e.g., 5%, 10%, 15%, 20%)
2. Use Data Table or Scenario Manager
3. Calculate profit for each scenario
4. Compare results and identify optimal discount rate

### Task 3: Linear Regression Analysis
**Objective:** Determine the relationship between Sales and Profit.

**Steps:**
1. Enable Data Analysis Toolpak (File > Options > Add-ins)
2. Select Data > Data Analysis > Regression
3. Set Profit as Y-variable (dependent)
4. Set Sales as X-variable (independent)
5. Interpret R-squared, coefficients, and p-values

### Task 4: Descriptive Statistics
**Objective:** Generate summary statistics for the dataset.

**Steps:**
1. Go to Data > Data Analysis > Descriptive Statistics
2. Select input range (sales, profit, quantity columns)
3. Check "Summary Statistics" and "Confidence Level"
4. Review mean, median, standard deviation, min, max values

### Task 5: Monthly Sales Growth Arrows
**Objective:** Add visual indicators for sales trends.

**Steps:**
1. Calculate month-over-month growth percentage
2. Apply custom number formatting
3. Use conditional formatting with custom icons:
   - Green up arrow (▲) for positive growth
   - Red down arrow (▼) for negative growth

**Formula Example:**
```excel
=(CurrentMonth - PreviousMonth) / PreviousMonth
```

### Task 6: Timestamp Creation
**Objective:** Create a dynamic timestamp column.

**Steps:**
1. Add a new column "Last Updated"
2. Use formula: `=NOW()`
3. Format as date-time (Ctrl+1 > Custom > "yyyy-mm-dd hh:mm:ss")

**Note:** NOW() updates automatically when the worksheet recalculates.

### Task 7: High-Value Customer Identification
**Objective:** Find customers with highest purchase values using advanced formulas.

**Steps:**
1. Group sales by customer using pivot or formulas
2. Use INDEX() to return customer names
3. Use MATCH() to find position of maximum value
4. Alternative: Use FILTER() or advanced filtering

**Formula Example:**
```excel
=INDEX(CustomerRange, MATCH(MAX(SalesRange), SalesRange, 0))
```

### Task 8: Pivot Table Analysis
**Objective:** Analyze total sales by region and product categories.

**Steps:**
1. Insert Pivot Table (Insert > PivotTable)
2. Add Region to Rows
3. Add Product to Columns
4. Add Sales to Values (Sum)
5. Format numbers as currency
6. Apply pivot table styles

### Task 9: Create Visualizations
**Objective:** Build charts to represent key performance indicators.

**Chart Types:**
- **Bar Chart:** Compare sales across regions or products
- **Line Chart:** Show sales trends over time (monthly/quarterly)
- **Pie Chart:** Display market share by product category or region

**Steps for Each Chart:**
1. Select data range
2. Insert > Choose chart type
3. Customize title, axis labels, and colors
4. Add data labels for clarity

### Task 10: Dashboard Summary
**Objective:** Create an executive dashboard with key insights.

**Dashboard Components:**
- KPI cards (Total Sales, Total Profit, Avg Order Value)
- Top 10 customers table
- Regional performance comparison (bar chart)
- Monthly trend analysis (line chart)
- Product mix (pie chart)
- Key insights and recommendations

---

## Worksheet Structure

The workbook should contain the following sheets:

1. **Raw Data** - Original dataset with sales transactions
2. **Conditional Formatting** - Top 10 customers highlighted
3. **What-If Analysis** - Discount scenario calculations
4. **Regression Analysis** - Output from Data Analysis Toolpak
5. **Descriptive Stats** - Statistical summary of key metrics
6. **Customer Analysis** - High-value customer identification
7. **Pivot Analysis** - Pivot tables for multi-dimensional analysis
8. **Charts** - Individual visualizations (bar, line, pie)
9. **Dashboard** - Executive summary with all KPIs and charts

---

## Required Excel Features

- Excel 2016 or later (recommended)
- Data Analysis Toolpak Add-in enabled
- Basic understanding of Excel formulas
- Familiarity with Pivot Tables and Charts

---

## Key Formulas Used

| Function | Purpose |
|----------|---------|
| `NOW()` | Create dynamic timestamp |
| `INDEX()` | Return value from array |
| `MATCH()` | Find position of value |
| `GROUPBY()` | Group and aggregate data (Excel 365) |
| `SUM()`, `AVERAGE()` | Basic aggregations |
| `IF()` | Conditional logic |
| `MAX()`, `MIN()` | Find extremes |

---

## Deliverables

Upon completion, the workbook should include:

✅ Conditional formatting applied to identify top performers  
✅ What-If Analysis scenarios documented  
✅ Linear Regression output with interpretation  
✅ Descriptive statistics summary  
✅ Growth indicators with arrows  
✅ Timestamp tracking implemented  
✅ High-value customers identified  
✅ Pivot tables with meaningful insights  
✅ Professional charts (bar, line, pie)  
✅ Executive dashboard with data storytelling  

---

## Expected Insights

By completing this analysis, you should be able to answer:

- Who are the top 10 most valuable customers?
- What is the optimal discount rate to maximize profit?
- How strong is the correlation between sales and profit?
- Which regions/products drive the most revenue?
- What are the sales trends over time?
- Which customer segments should be prioritized for retention?

---

## Tips for Success

1. **Data Validation:** Ensure data is clean before analysis
2. **Consistent Formatting:** Use uniform number formats and styles
3. **Clear Labels:** Name ranges and sheets descriptively
4. **Documentation:** Add notes explaining complex formulas
5. **Color Coding:** Use consistent color schemes across charts
6. **Error Handling:** Use IFERROR() to manage calculation errors
7. **Refresh Data:** Update pivot tables after data changes
8. **Professional Layout:** Align elements and use white space effectively

---

## Learning Outcomes

After completing this project, you will be able to:

- Apply advanced Excel functions for data analysis
- Perform statistical analysis using built-in tools
- Create compelling visualizations and dashboards
- Use conditional formatting for data insights
- Conduct scenario analysis for decision-making
- Extract actionable insights from raw data
- Present findings in a business-ready format

---

## Next Steps

1. Prepare your dataset (sales, customers, products, regions, dates)
2. Follow each task sequentially
3. Document your findings in each sheet
4. Build the final dashboard
5. Present insights and recommendations

---

## Support & Resources

- Excel Data Analysis Toolpak Documentation
- Pivot Table Best Practices
- Dashboard Design Principles
- Statistical Analysis in Excel

---

**Project Status:** Ready for Implementation  
**Estimated Completion Time:** 4-6 hours  
**Difficulty Level:** Intermediate to Advanced

---

*This project demonstrates practical Excel skills essential for business analysts, data analysts, and financial professionals.*
