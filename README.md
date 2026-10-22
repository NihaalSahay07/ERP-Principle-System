# Inventory Management System | Excel-Based ERP

An inventory management workbook built in **Microsoft Excel** to bring purchases, sales, stock tracking, and business reporting into one place.

**Author:** [Nihaal Sahay](https://github.com/NihaalSahay07)  
**Project period:** February–March 2025  
**Tools:** Microsoft Excel, Excel Tables, formulas, PivotTables, PivotCharts, and data validation

## About the project

Keeping track of stock becomes difficult when purchase records, sales records, and product information are maintained separately. Even a small mistake in one sheet can lead to incorrect stock figures or time-consuming manual checks.

I built this project to explore how the core ideas behind an Enterprise Resource Planning (ERP) system could work inside Excel. Instead of updating inventory and reports by hand after each transaction, the workbook connects the underlying records so that the relevant calculations and summaries can update from the data entered.

It is an **ERP-inspired Excel solution**, not a replacement for a full ERP platform. The focus is on making everyday inventory operations easier to follow and manage, especially in a small-business setting.

## What the workbook does

- **Tracks purchases and sales:** Records incoming and outgoing stock in separate transaction sheets.
- **Calculates stock levels:** Uses transaction data to work out the quantity currently available for each product.
- **Keeps reference data organized:** Maintains product, customer, and vendor details in dedicated sheets.
- **Summarizes business activity:** Brings together sales, purchases, estimated profit, inventory value, and product performance in a dashboard.
- **Flags low stock:** Uses conditional formatting to highlight items that may need reordering.
- **Makes data entry easier:** Uses dropdowns and quick-entry navigation to reduce repetitive work.

## Workbook structure

The system is organized into eight connected modules:

| Module | Purpose |
| --- | --- |
| **Dashboard** | Presents sales, purchase, profit, stock, and customer summaries. |
| **Product Master** | Stores product names, HSN codes, categories, cost prices, and selling prices. |
| **Customer Database** | Maintains customer names, contact details, and locations. |
| **Vendor Database** | Maintains supplier details, locations, and product mappings. |
| **Purchase Module** | Records stock received from suppliers. |
| **Sales Module** | Records sales and the quantities sold. |
| **Inventory Module** | Calculates available stock from purchase and sales records. |
| **New Entry Module** | Provides shortcuts for adding purchases and sales. |

### How the parts work together

```text
                 Product Master
                       |
          +------------+------------+
          |                         |
     Purchase Records          Sales Records
          |                         |
          +------------+------------+
                       |
               Stock Calculation
                       |
               Dashboard & Reports
```

For example, when a purchase is added, the purchased quantity contributes to the stock calculation. When a sale is recorded, the sold quantity is subtracted. The dashboard can then use those records to display the latest calculated totals.

## Project figures

The original project documentation reports the following figures:

| Metric | Reported value |
| --- | ---: |
| Inventory managed | ₹29.84 lakh |
| Transaction value | ₹85+ lakh |
| Purchases | ₹55.4 lakh |
| Sales | ₹30.7 lakh |
| Profit | ₹5.11 lakh |
| Product categories | 10 |
| Active customers | 7 |
| Inventory accuracy | 99.8% |

These are **project-reported figures**, not independently audited results. They describe the workbook's documented dataset and outcomes.

## How inventory is calculated

The main stock calculation follows a simple rule:

**Available stock = Total units purchased − Total units sold**

An example Excel formula using structured table references is:

```excel
=SUMIF(Purchase[HSN_Code],[@HSN_Code],Purchase[Units])-SUMIF(Sales[HSN_Code],[@HSN_Code],Sales[Units])
```

The formula matches records using each product's HSN code, adds the purchased units, and subtracts the sold units.

Inventory value can then be calculated using the available quantity and the stored unit cost:

```excel
=[@Stock]*[@Cost]
```

The project documentation also discusses FIFO, weighted average cost, and market-value approaches to inventory valuation. Those approaches require their own costing rules; the simple formula above illustrates stock quantities and a basic cost-based valuation, rather than implementing all three methods.

## Excel features used

| Excel feature | How it is used |
| --- | --- |
| Excel Tables | Keeps transaction and master data structured and expandable. |
| `SUMIF` / `SUMIFS` | Aggregates purchases, sales, and stock-related figures. |
| `VLOOKUP` | Retrieves product information from reference tables. |
| `INDEX` + `MATCH` | Supports lookups across customer and vendor records. |
| Data Validation | Creates dropdown options and helps limit invalid entries. |
| `IFERROR` | Handles lookup failures more clearly. |
| PivotTables | Groups and summarizes business data. |
| PivotCharts | Visualizes trends and comparisons on the dashboard. |
| Conditional Formatting | Highlights stock levels and exceptions requiring attention. |

## Dashboard and reporting

The dashboard is intended to provide a quick picture of the business without having to inspect every transaction sheet. It brings together:

- Sales, purchases, and profit
- Total inventory value and product count
- Active customers and top customers
- Best-selling products
- Stock-level analysis and reorder alerts

One useful example is a low-stock warning: when a product falls below its reorder threshold, the workbook can flag it so that the item is easier to spot during review.

## What I learned

This project helped me connect inventory management concepts with a practical spreadsheet workflow. The most valuable part was designing the relationship between different sheets: an entry in one part of the workbook should flow through to the stock calculations and business summaries without requiring the same information to be entered repeatedly.

It also gave me hands-on experience with structured data, lookup and aggregation formulas, dashboard design, basic reporting, and the importance of keeping business records consistent.

The original project documentation reports a **75% reduction in manual work** and **99.8% inventory accuracy**. Those are reported outcomes rather than measurements independently validated in this README.

## Getting started

1. Download or clone this repository.
2. Open [`INVENTORY MANAGEMENT IN EXCEL.xlsx`](./INVENTORY%20MANAGEMENT%20IN%20EXCEL.xlsx) in Microsoft Excel.
3. Explore the **Product Master**, **Customer Database**, and **Vendor Database** to understand the reference data.
4. Review the **Purchase** and **Sales** sheets to see how transactions are recorded.
5. Open the **Inventory** and **Dashboard** sheets to explore the calculated results and reports.

> **Tip:** Work on a copy of the workbook if you want to experiment with the data. Depending on the Excel version and workbook settings, you may need to refresh PivotTables or related reports after changing transaction records.

## Possible improvements

There are several directions in which I would like to extend this project:

- Barcode-based product entry
- More streamlined data entry with VBA
- Automated email reports
- Support for stock across multiple warehouses
- A database backend, such as SQL, for larger datasets
- Power BI reporting
- Cloud access and integration with dedicated ERP systems

These are ideas for future development, not features claimed to be available in the current workbook.

## Repository contents

```text
ERP-Principle-System/
├── INVENTORY MANAGEMENT IN EXCEL.xlsx
└── README.md
```

## Author

**Nihaal Sahay**  
GitHub: [@NihaalSahay07](https://github.com/NihaalSahay07)

---

*Built as a hands-on exploration of how ERP principles and inventory reporting can be applied using Microsoft Excel.*
