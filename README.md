# 📊 PR.2 Analyzer

> An Excel project that turns raw sales data into answers and insights using conditional formatting, What-If analysis, regression, descriptive statistics, pivot tables, charts and a dashboard.

---

## 🎥 Project Experience Video

▶️ **[Click here to watch the project video](https://drive.google.com/file/d/1utP_lzRpINAYptlzaqpcgD8XOjn53Y0B/view?usp=sharing)**

In this video I explain every sheet, show how each task works and share what I learned.

---

## 📑 Table of Contents

- [Project Overview](#-project-overview)
- [Objectives](#-objectives)
- [Sheets and Data](#️-sheets-and-data)
- [Tools Used](#️-tools-used)
- [Tasks Performed](#-tasks-performed)
- [Data Story and Insights](#-data-story-and-insights)
- [Key Observations](#-key-observations)
- [Assumptions and Limitations](#️-assumptions--limitations)
- [Project Structure](#-project-structure)
- [How to Open the Project](#️-how-to-open-the-project)
- [What I Learned](#-what-i-learned)
- [Conclusion](#-conclusion)

---

## 📌 Project Overview

Raw data on its own does not help a business. It becomes useful only when it is sorted, measured, compared and shown in a clear way.

**PR.2 Analyzer** is an Excel analysis project built on one sales dataset. The work is split into 10 tasks, and every task has its own sheet. Most results come from formulas, so they update by themselves when the data changes.

The project includes:

- 🎨 Conditional Formatting (Top 10 highlight)
- 🔮 What-If Analysis (discount vs profit)
- 📈 Linear Regression (Profit vs Sales)
- 🧮 Descriptive Statistics and Histogram
- 🔼 Up / Down arrows with custom formatting
- 🕒 Timestamp with `NOW()`
- 🏆 High-Value Customers with `INDEX` + `MATCH` and AutoFilter
- 🧊 Pivot Table (Region × Product)
- 📊 Bar, Line and Pie charts
- 🖥️ Dashboard with insights (Data Storytelling)

---

## 🎯 Objectives

1. Highlight the top 10 customers by total purchase using conditional formatting.
2. Study how a change in discount changes the total profit (What-If Analysis).
3. Find the relationship between Sales and Profit with Linear Regression.
4. Summarise the data with Descriptive Statistics and a Histogram.
5. Show monthly sales growth with up and down arrows using a custom number format.
6. Create a timestamp column using `NOW()`.
7. Identify high-value customers using `INDEX`, `MATCH` and filters.
8. Build a Pivot Table of total sales by region and product.
9. Visualise KPIs with bar, line and pie charts.
10. Summarise the key insights in a dashboard.

---

## 🗂️ Sheets and Data

The workbook has **16 sheets**. `Data` is the main dataset and every other sheet is built from it.

| Sheet | Purpose |
|-------|---------|
| 🏠 **Start_Here** | Guide to all sheets |
| 🗃️ **Data** | Main dataset (Order ID, Date, Customer, Region, Product, Sales, Cost, Profit, Month No.) |
| 🎨 **Cond_Formatting** | Task 1: Top 10 customers |
| 🔮 **What_If** | Task 2: Discount vs profit |
| 📈 **Regression** | Task 3: Regression using formulas |
| 📈 **Regression_Toolpak** | Task 3: Regression in Data Analysis Toolpak layout |
| 🧮 **Desc_Stats** | Task 4: Descriptive statistics using formulas, plus histogram |
| 🧮 **DescStats_Toolpak** | Task 4: Descriptive Statistics in Toolpak layout |
| 📊 **Histogram_Toolpak** | Task 4: Histogram table and chart in Toolpak layout |
| 🔼 **Growth** | Task 5: Monthly growth with arrows |
| 🕒 **Timestamp** | Task 6: `NOW()` timestamp |
| 🏆 **High_Value** | Task 7: High-value customers with AutoFilter |
| 🧊 **Pivot_Table** | Task 8: Real Excel Pivot Table |
| 🧱 **Chart_Data** | Helper table that feeds the charts |
| 📊 **Charts** | Task 9: Bar, line and pie charts |
| 🖥️ **Dashboard** | Task 10: KPIs, insights and charts |

**About the dataset**

| Item | Details |
|------|---------|
| Orders | 60 sample orders |
| Customers | 15 |
| Regions | North, South, East, West |
| Products | Laptop, Mobile, Tablet, Headphones, Printer |
| Profit rule | Profit = Sales − Cost |

---

## 🛠️ Tools Used

| Tool | Purpose |
|------|---------|
| 📗 Microsoft Excel | Formulas, Pivot Table, charts and conditional formatting |
| 🧰 Data Analysis Toolpak | Regression, Descriptive Statistics and Histogram layout |
| ☁️ Google Drive | Storing the project video |
| 🐙 GitHub | Hosting the project |

---

## 🔍 Tasks Performed

### 🎨 Task 1: Top 10 Customers (`Cond_Formatting`)

| Step | Formula / Feature | Purpose |
|------|-------------------|---------|
| Total purchase per customer | `=SUMIF(Data!C:C, A4, Data!F:F)` | Adds all sales of one customer |
| Rank of each customer | `=RANK(B4,$B$4:$B$18)` | Gives rank 1, 2, 3 ... |
| Highlight rule | `=$B4>=LARGE($B$4:$B$18,10)` | Turns the top 10 rows green |

### 🔮 Task 2: What-If Analysis (`What_If`)

| Step | Formula / Feature | Purpose |
|------|-------------------|---------|
| Discount input | Yellow input cell | Change it to test any discount |
| Discount amount | `=B4*B6` | Sales × discount % |
| Net sales | `=B4-B7` | Sales after discount |
| Total profit | `=B8-B5` | Net sales − total cost |
| Discount table | `=$B$4*(1-A13)` and `=B13-$B$5` | Shows 0% to 30% discount at once |

### 📈 Task 3: Linear Regression (`Regression`, `Regression_Toolpak`)

| Step | Formula / Feature | Purpose |
|------|-------------------|---------|
| Settings | Y = Profit, X = Sales | Profit depends on Sales |
| Fit values | `=RSQ()`, `=STEYX()`, `=LINEST()` | R Square, Standard Error, coefficients |
| ANOVA | `=FDIST()` | Significance F |
| Coefficient tests | `=TDIST()`, `=TINV()` | P-value and 95% limits |
| Scatter chart | Chart with trend line | Shows the relationship visually |

### 🧮 Task 4: Descriptive Statistics and Histogram (`Desc_Stats`, `DescStats_Toolpak`, `Histogram_Toolpak`)

| Step | Formula / Feature | Purpose |
|------|-------------------|---------|
| Centre of data | `=AVERAGE()`, `=MEDIAN()`, `=MODE()` | Mean, Median, Mode |
| Spread of data | `=STDEV()`, `=VAR()`, `=MAX()-MIN()` | Standard Deviation, Variance, Range |
| Shape of data | `=KURT()`, `=SKEW()` | Kurtosis and Skewness |
| Histogram table | `=COUNTIFS()` with bins | Frequency of each Sales range |
| Histogram chart | Column chart | Shows how Sales are spread |

### 🔼 Task 5: Up and Down Arrows (`Growth`)

| Step | Formula / Feature | Purpose |
|------|-------------------|---------|
| Monthly sales | `=SUMIF(Data!I:I, A4, Data!F:F)` | Total sales of each month |
| Growth % | `=(C5-C4)/C4` | Change from the previous month |
| Arrow format | Custom number format with ▲ and ▼ | Shows ▲ for growth and ▼ for fall |
| Up / Down text | `=IF(D5>0,"▲ Up",IF(D5<0,"▼ Down","■ Same"))` | Writes the direction |
| Colours | Conditional formatting | Green for Up, red for Down |

### 🕒 Task 6: Timestamp (`Timestamp`)

| Step | Formula / Feature | Purpose |
|------|-------------------|---------|
| Current date and time | `=NOW()` | Live timestamp |
| Today's date | `=TODAY()` | Date only |
| Timestamp column | `=IF(A8<>"",NOW(),"")` | Adds a timestamp next to each order |

### 🏆 Task 7: High-Value Customers (`High_Value`)

| Step | Formula / Feature | Purpose |
|------|-------------------|---------|
| Total purchase | `=SUMIF()` | Total of each customer |
| Number of orders | `=COUNTIF()` | Orders per customer |
| Status | `=IF(B4>=$B$2,"High Value","Regular")` | Compares with the limit |
| Top customer | `=INDEX(A4:A18,MATCH(MAX(B4:B18),B4:B18,0))` | Name of the biggest buyer |
| Top 3 customers | `=LARGE()` with `INDEX` + `MATCH` | Best three customers |
| Filter | AutoFilter on Status | Shows only High Value customers |

### 🧊 Task 8: Pivot Table (`Pivot_Table`)

| Setting | Value |
|---------|-------|
| Rows | Region |
| Columns | Product |
| Values | Sum of Sales |
| Result | Total sales of each product in each region, with Grand Total |

`Chart_Data` holds the same numbers with `SUMIFS`, so the charts always have fixed cells to read.

### 📊 Task 9: Charts (`Charts`)

| Chart | Data Shown | Best For |
|-------|-----------|----------|
| 📊 Bar chart | Sales by region | Comparing regions |
| 📈 Line chart | Sales by month | Seeing the trend |
| 🥧 Pie chart | Sales share by product | Seeing each product's share |

### 🖥️ Task 10: Dashboard (`Dashboard`)

| Part | Content |
|------|---------|
| KPIs | Total Sales, Total Profit, Profit Margin, Total Orders, Top Customer, High-Value Customers, Best Region, Best Product, Best Month |
| Insights | 5 lines of text built with formulas, so they update with the data |
| Charts | Bar, line and pie charts |

---

## 📖 Data Story and Insights

The dashboard tells a simple story:

1. **How big is the business?** Total sales, total profit and the profit margin are shown first.
2. **Where does it sell best?** The best region, best product and best month are picked out with `INDEX` + `MATCH`.
3. **Who buys the most?** The top customer and the High Value customers are listed.
4. **What happens if we change the discount?** The What-If sheet shows that more discount means less profit.
5. **How are Sales and Profit linked?** Regression shows that higher sales bring higher profit.

---

## 💡 Key Observations

- A higher discount lowers the profit, and a very high discount can cause a loss.
- Sales and Profit move together, so the regression line slopes upward.
- `LARGE` inside a conditional formatting rule keeps the top 10 highlight correct when data changes.
- `INDEX` + `MATCH` finds the name of the top customer without any manual search.
- A custom number format can show ▲ and ▼ without a second column.
- `NOW()` is a volatile function, so it updates whenever the sheet calculates.
- A Pivot Table gives a Region × Product summary in a few clicks, and it must be refreshed after the data changes.
- Bar, line and pie charts answer different questions: comparison, trend and share.

---

## ⚠️ Assumptions & Limitations

**Assumptions**

- The sample data was created for practice so that every task gives a clear result.
- Profit is calculated as Sales − Cost.
- In the What-If sheet, the discount is applied on total sales, and cost stays the same.
- The High-Value limit is kept in one yellow cell, so changing it updates the status of every customer.

**Limitations**

- `Regression_Toolpak`, `DescStats_Toolpak` and `Histogram_Toolpak` follow the layout of Data Analysis Toolpak output and hold fixed values. The formula sheets `Regression` and `Desc_Stats` are the ones that update with the data.
- The Pivot Table must be refreshed (right-click → Refresh) after the data changes.
- The AutoFilter hides regular customers with fixed hidden rows, so clear the filter and apply it again if the limit changes.
- Some previews (like phone file viewers) may show formula cells as empty. Open the file in Excel to see all results.
- The data is small and made for practice, not real business data.

---

## 📁 Project Structure

```
PR2-Analyzer/
│
├── 📄 PR2_Analyzer.xlsx   → 16 sheets with data, formulas, Pivot Table, charts and dashboard
└── 📄 README.md           → project documentation
```

---

## ▶️ How to Open the Project

1. Download or clone the repository.
2. Open **Microsoft Excel**.
3. Open the file `PR2_Analyzer.xlsx`.
4. Start from the `Start_Here` sheet and go through the sheets in order.
5. Change a yellow cell (for example the discount in `What_If`) and watch the results update.
6. To use your own data, paste it in the `Data` sheet (columns A to G), then refresh the Pivot Table.

---

## 📚 What I Learned

- How conditional formatting can highlight important rows automatically.
- How What-If analysis helps in choosing a better discount.
- How regression shows the relationship between two numbers.
- How descriptive statistics and a histogram describe a whole dataset.
- How custom formatting adds arrows and colours to numbers.
- How `INDEX` + `MATCH` finds a value without searching by hand.
- How a Pivot Table summarises data from two sides.
- How a dashboard puts the main results on one page.

One main lesson: **good analysis is not about having more data, it is about showing the right answer clearly.**

---

## 📌 Conclusion

This project shows how raw sales data can become clear decisions. It covers the full flow:

**Highlight → Test (What-If) → Predict (Regression) → Summarise → Track Growth → Find Top Customers → Pivot → Visualise → Dashboard**

---

## 👨‍💻 Author

**Shyam Rajgor**

Excel Learner

Currently learning Excel and data analysis and building practical skills through hands-on projects.

---

## 🙏 Acknowledgement

This project was created for learning purposes, based on the PR.2 Analyzer project assignment.

---

## ⭐ Project Summary

| | |
|---|---|
| **Project** | PR.2 Analyzer |
| **Author** | Shyam Rajgor |
| **Tool** | Microsoft Excel |
| **Total Sheets** | 16 |
| **Total Tasks** | 10 |
| **Topics** | Conditional Formatting • What-If • Regression • Descriptive Statistics • Histogram • Arrows • Timestamp • INDEX-MATCH • Pivot Table • Charts • Dashboard |
| **Video** | [Watch on Google Drive](https://drive.google.com/file/d/1utP_lzRpINAYptlzaqpcgD8XOjn53Y0B/view?usp=sharing) |

*Thank you for reviewing my project.*