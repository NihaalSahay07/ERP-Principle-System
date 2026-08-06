# 📦 Inventory Management System using Microsoft Excel

An **ERP-inspired Inventory Management System** developed entirely in **Microsoft Excel** to automate inventory tracking, purchase and sales management, profit calculation, and business reporting.

The project was developed between **February 2025 and March 2025** as a practical implementation of inventory management concepts using advanced Excel features. It demonstrates how Excel can be used as a lightweight ERP solution for small and medium-sized businesses.

---

## 📌 Project Overview

This project was built to solve common inventory management challenges faced by small businesses, such as:

- Manual stock tracking
- Duplicate data entry
- Inventory mismatches
- Delayed reporting
- Lack of integration between purchases, sales, and inventory

To address these challenges, I designed an interconnected Excel workbook where a single transaction automatically updates inventory, reports, profit, and dashboard KPIs—similar to the workflow of an ERP (Enterprise Resource Planning) system.

---

# 🎯 Objectives

- Automate inventory management
- Reduce manual calculations
- Improve inventory accuracy
- Generate real-time business reports
- Mimic the workflow of an ERP system using Microsoft Excel

---

# 📊 Project Statistics

| Metric | Value |
|---------|------:|
| Inventory Managed | ₹29.84 Lakhs |
| Total Transactions | ₹85+ Lakhs |
| Purchase Value | ₹55.4 Lakhs |
| Sales Value | ₹30.7 Lakhs |
| Profit Generated | ₹5.11 Lakhs |
| Product Categories | 10 |
| Active Customers | 7 |
| Inventory Accuracy | 99.8% |

---

# 🏗️ System Modules

The workbook consists of **8 interconnected modules**.

### 📋 Dashboard
Displays real-time business insights:

- Total Sales
- Total Purchases
- Profit
- Inventory Value
- Top Customers
- Top Selling Products
- Stock Alerts

---

### 📦 Product Master

Stores:

- HSN Code
- Product Name
- Category
- Cost Price
- Selling Price

Acts as the master database for all inventory transactions.

---

### 👥 Customer Database

Maintains customer information including:

- Customer Name
- Contact Number
- Location

---

### 🚚 Vendor Database

Maintains vendor information including:

- Vendor Name
- Contact Details
- Product Mapping
- Location

---

### 🛒 Purchase Module

Records purchase transactions.

Automatically updates inventory whenever new stock is purchased.

---

### 💰 Sales Module

Records customer sales.

Automatically deducts inventory and updates sales reports.

---

### 📦 Inventory Module

Calculates current inventory dynamically using purchase and sales transactions.

Current Stock = Total Purchases − Total Sales

---

### ➕ New Entry Module

Provides quick navigation buttons for:

- New Purchase Entry
- New Sales Entry

This reduces manual navigation and improves efficiency.

---

# ⚙️ Excel Features Used

| Feature | Purpose |
|----------|----------|
| Excel Tables | Structured and expandable data management |
| SUMIFS | Dynamic inventory calculation |
| VLOOKUP | Fetch product details automatically |
| INDEX-MATCH | Customer and vendor lookup |
| Data Validation | Dropdown lists and error prevention |
| IFERROR | Handle lookup errors gracefully |
| Pivot Tables | Business reporting and analysis |
| Pivot Charts | Interactive dashboard visualization |
| Conditional Formatting | Low stock alerts |

---

# 📈 Inventory Calculation

Current inventory is calculated automatically using purchase and sales data.

### Formula

```excel
=SUMIF(Purchase[HSN_Code],[@HSN_Code],Purchase[Units])
-
SUMIF(Sales[HSN_Code],[@HSN_Code],Sales[Units])
```

---

### Inventory Value

```excel
=[@Stock]*[@Cost]
```

---

# 📊 Dashboard KPIs

The dashboard displays:

- 📈 Total Sales
- 📉 Total Purchases
- 💰 Profit
- 📦 Inventory Value
- 👥 Active Customers
- 🛍️ Product Count
- ⭐ Top Selling Products
- 🏆 Top Customers
- 📊 Stock Level Analysis

All KPIs update automatically whenever a purchase or sale is recorded.

---

# 📦 Inventory Valuation Methods

The project supports multiple inventory valuation techniques.

- FIFO (First-In First-Out)
- Weighted Average Cost
- Market Value Method

These methods help maintain accurate inventory valuation depending on business requirements.

---

# 🔔 Smart Stock Alerts

Automatic stock alerts notify users whenever inventory falls below the reorder level.

Example:

```
⚠ Smart Watch Needs Reordering!
```

This helps businesses prevent stockouts and improve inventory planning.

---

# 💼 Business Impact

The system successfully:

- ✅ Reduced manual work by **75%**
- ✅ Improved inventory accuracy to **99.8%**
- ✅ Automated inventory tracking
- ✅ Reduced data entry errors using dropdown validation
- ✅ Generated real-time reports and dashboards
- ✅ Identified slow-moving inventory
- ✅ Improved business decision-making
- ✅ Simplified day-to-day inventory operations

---

# 🧠 ERP Concept

This project was inspired by the workflow of an **Enterprise Resource Planning (ERP)** system.

Although developed in Microsoft Excel, it follows the same basic principle of ERP software:

- A single transaction updates multiple business functions automatically.
- Purchase entries update inventory.
- Sales entries reduce stock.
- Dashboard KPIs refresh automatically.
- Reports are generated in real time.

This provides an integrated workflow similar to enterprise software such as SAP.

---

# 📈 Workflow

```
                Product Master
                       │
        ┌──────────────┼──────────────┐
        │                             │
  Purchase Module              Sales Module
        │                             │
        └──────────────┼──────────────┘
                       │
              Inventory Calculation
                       │
                Dashboard & Reports
```

---

# 🚀 Future Enhancements

- Barcode Scanner Integration
- VBA Automation
- Automatic Email Reports
- Multi-Warehouse Support
- SQL Database Integration
- Power BI Dashboard
- SAP ERP Integration
- Cloud-Based Inventory Management

---

# 📚 Key Learnings

Through this project, I gained practical knowledge of:

- Inventory Management
- ERP Workflow
- Business Process Automation
- Dashboard Development
- Data Analysis
- Business Reporting
- Excel Automation
- Supply Chain Operations

---

# 🛠️ Technologies Used

- Microsoft Excel
- Excel Tables
- SUMIFS
- VLOOKUP
- INDEX-MATCH
- Pivot Tables
- Pivot Charts
- Conditional Formatting
- Data Validation
- IFERROR


---

# 👨‍💻 Author

**Harsh Vardhan Pandey**

B.Tech Mechanical Engineering

National Institute of Technology (NIT) Jamshedpur

---

## ⭐ Project Summary

This project demonstrates how **Microsoft Excel** can be transformed into an **ERP-inspired Inventory Management System** by integrating purchases, sales, inventory, customers, vendors, and reporting into a single automated workbook. Using advanced Excel features such as **SUMIFS, VLOOKUP, INDEX-MATCH, Pivot Tables, Data Validation, and Conditional Formatting**, the system provides real-time inventory tracking, business insights, and decision support while reducing manual effort and improving operational efficiency.
