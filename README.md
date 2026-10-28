# 📊 Excel Data Exploration: Product Database Case Study

> Turning raw product data into clear insights using Excel formulas.

**Author:** Hari Hara Kumar
**Source:** Entri Academy, Data Analytics Module 1
**Tool:** Microsoft Excel
**Level:** Beginner

---

## 🎯 Objective
Explore a product database to summarise key numbers, classify products by price, and extract structured information from Product IDs, using built-in Excel functions.

## 🗂️ Dataset
The data is stored as an Excel Table named `Product_DB`.

| Column | Description |
|---|---|
| Product ID | Coded ID containing day, month and country |
| Product Name | Name of the product |
| Brand Name | Brand of the product |
| Quantity | Units available |
| Category | Product category |
| Price ($) | Price in USD |

---

## ❓ Questions and Answers

### 1️⃣ Sum, Count and Average

**Q1. Calculate the total price of all products in the dataset.**
**Answer: 10,100**
```excel
=SUM(Product_DB[Price ($)])
```

**Q2. Determine how many products exist in the dataset.**
**Answer: 34 total records (21 unique after removing duplicates)**
```excel
=COUNT(Product_DB[Price ($)])
```
Unique count: Data → Remove Duplicates

**Q3. Calculate the average price of the products.**
**Answer: 297.06**
```excel
=AVERAGE(Product_DB[Price ($)])
```

### 2️⃣ Minimum and Maximum Values

**Q4. Identify the minimum price among all products.**
**Answer: 30**
```excel
=MIN(Product_DB[Price ($)])
```

**Q5. Identify the maximum price among all products.**
**Answer: 1,000**
```excel
=MAX(Product_DB[Price ($)])
```

### 3️⃣ Logical Function: IF

**Q6. Create a new column named Price Range and categorise products using IF.**
- If Price ≥ $500 → **High Price**
- If Price < $500 → **Standard Price**

**Answer:** New column **Price Range** created. Result: **7 High Price** and **27 Standard Price** products.
```excel
=IF([@[Price ($)]]>=500,"High Price","Standard Price")
```

### 4️⃣ Conditional Functions: SUMIF and COUNTIF

**Q7. Use SUMIF to calculate the total price of products in the Electronics category.**
**Answer: 8,050**
```excel
=SUMIF(Product_DB[Category],"Electronics",Product_DB[Price ($)])
```

**Q8. Use COUNTIF to determine the number of products with a price less than $100.**
**Answer: 11**
```excel
=COUNTIF(Product_DB[Price ($)],"<100")
```

### 5️⃣ Text Functions: LEFT, RIGHT, MID

**Q9. Create a column named Day using the first 2 characters of Product ID (LEFT).**
**Answer:** New column **Day** created.
```excel
=LEFT([@[Product ID]],2)
```

**Q10. Create a column named Country Code using the last 2 characters of Product ID (RIGHT).**
**Answer:** New column **Country Code** created. There are 12 unique codes, with **CA** appearing most often.
```excel
=RIGHT([@[Product ID]],2)
```

**Q11. Create a column named Month using the 4th to 6th characters of Product ID (MID).**
**Answer:** New column **Month** created. Values span January to December.
```excel
=MID([@[Product ID]],4,3)
```

---

## 📋 Summary of Results

| # | Question | Answer | Function |
|---|---|---|---|
| 1 | Total price | **10,100** | `SUM` |
| 2 | Number of products | **34** (21 unique) | `COUNT` |
| 3 | Average price | **297.06** | `AVERAGE` |
| 4 | Minimum price | **30** | `MIN` |
| 5 | Maximum price | **1,000** | `MAX` |
| 6 | Price Range | **7 High / 27 Standard** | `IF` |
| 7 | Electronics total | **8,050** | `SUMIF` |
| 8 | Products under $100 | **11** | `COUNTIF` |
| 9 | Day | New column | `LEFT` |
| 10 | Country Code | New column | `RIGHT` |
| 11 | Month | New column | `MID` |

## 🔍 Key Insights
- **Price distribution:** Most products sit below $500, with only 7 of 34 in the High Price band.
- **Electronics dominates value:** Electronics contributes 8,050 of the 10,100 total price (about 80%).
- **Budget-friendly range:** 11 products cost under $100.
- **Product ID patterns:** Day and month values span January to December, with 12 unique country codes.
- **Data quality note:** 34 total records reduce to 21 after removing duplicates, so totals and averages include repeated entries.

## 📦 Deliverables
- ✅ Excel workbook (.xlsx) with formulas, results and new columns
- ✅ PDF with screenshots of results, new columns and formula-bar views
- ✅ This README

## 🚀 Skills Practised
- Data summarisation (SUM, COUNT, AVERAGE, MIN, MAX)
- Logical classification (IF)
- Conditional aggregation (SUMIF, COUNTIF)
- Text extraction (LEFT, MID, RIGHT)
- Structured references with Excel Tables
- Building clean, clearly labelled worksheets

## ✅ Conclusion
This case study builds a beginner-level foundation in Excel-based data exploration, from basic statistics to conditional logic and text manipulation. These skills support product reporting, data checking and business analysis.
