# 📊 Excel Data Exploration: Product Dataset
Excel data exploration project: summarize, classify and decode a product dataset using SUM, AVERAGE, MIN/MAX, IF, SUMIF, COUNTIF, LEFT, MID &amp; RIGHT. Beginner-friendly, formula-driven analysis.
> Turning raw product data into clear insights, using nothing but Excel formulas.

**Author:** Hari Hara Kumar
**Source:** Entri Academy, Data Analytics Module 1
**Tool:** Microsoft Excel
**Level:** Beginner

---

## 🎯 Objective
Explore a product dataset to summarize key numbers, classify products by price, and pull structured information out of Product IDs, all with built-in Excel functions.

## 🗂️ Dataset
| Column | Description |
|---|---|
| Product ID | Coded ID containing day, month and country |
| Product Name | Name of the product |
| Brand Name | Brand of the product |
| Quantity | Units available |
| Category | Product category |
| Price | Price in USD |

## 🧰 Functions Used
| Task | Function(s) |
|---|---|
| Total, count, average price | `SUM`, `COUNT`/`COUNTA`, `AVERAGE` |
| Lowest and highest price | `MIN`, `MAX` |
| Price classification | `IF` |
| Category total, low-price count | `SUMIF`, `COUNTIF` |
| Extract Day, Month, Country Code | `LEFT`, `MID`, `RIGHT` |

### Sample formulas (assuming Product ID in A, Category in E, Price in F)
```excel
=IF(F2>=500,"High Price","Standard Price")
=SUMIF(E:E,"Electronics",F:F)
=COUNTIF(F:F,"<100")
=LEFT(A2,2)      → Day
=MID(A2,4,3)     → Month
=RIGHT(A2,2)     → Country Code
```

## 🔍 Key Findings
| Metric | Result |
|---|---|
| Total price of all products | **10,100** |
| Number of products | **34** |
| Average price | **297.06** |
| Minimum price | **30** |
| Maximum price | **1,000** |
| Electronics total price | **8,050** |
| Products priced below $100 | **11** |

**💡 Price split:** 7 products are *High Price* (≥ $500) and 27 are *Standard Price* (< $500). Most of the catalogue sits below the $500 mark.

**🆔 Product ID insights:** The extracted day and month values span January to December. There are 12 unique country codes, with **CA****CANADA** appearing most often.

## 📦 Deliverables
- ✅ Excel workbook (.xlsx) with formulas, results and new columns
- ✅ PDF with screenshots of results, new columns and formula-bar views
- ✅ This README

## 🚀 Skills Practised
- Summarising numerical data
- Classifying records with logical conditions
- Conditional aggregation (SUMIF / COUNTIF)
- Text extraction from structured IDs
- Building clean, clearly labelled worksheets

## ✅ Conclusion
A hands-on foundation in Excel-based data exploration: from basic statistics to logic, conditional aggregation and text manipulation. These are the building blocks for product reporting, data checking and business analysis.
