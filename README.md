# 🧹📊 Excel - Product Sales Data Cleaning & Imputation

> An Excel-based Data Cleaning project that turns a messy product sales dataset into a clean, analysis-ready table using Excel formulas, imputation and standardisation.

---

## 📑 Table of Contents

- [🎯 Project Overview](#-project-overview)
- [📁 Repository Structure](#-repository-structure)
- [🗂️ Workbook Sheets](#️-workbook-sheets)
- [🧾 Data Dictionary](#-data-dictionary)
- [🚨 Data Quality Issues Found](#-data-quality-issues-found)
- [🛠️ Cleaning & Imputation Steps](#️-cleaning--imputation-steps)
- [🧮 Excel Formulas Used](#-excel-formulas-used)
- [🔄 Before vs After](#-before-vs-after)
- [📈 Final Dataset Summary](#-final-dataset-summary)
- [🚀 How to Use](#-how-to-use)
- [💡 Key Learnings](#-key-learnings)
- [🔮 Future Improvements](#-future-improvements)
- [👤 Author](#-author)
- [📚 Acknowledgement](#-acknowledgement)

---

## 🎯 Project Overview

This project cleans and prepares a **Product Sales dataset** covering electronics, fashion, kitchen, outdoor and accessory products sold across multiple countries.

The raw data contains missing values, inconsistent text, truncated categories and duplicate records. The workbook walks through each cleaning stage, from raw data to the final cleaned table.

**🎯 Goals**

- ✅ Remove duplicate records
- ✅ Fix inconsistent text casing and truncated values
- ✅ Impute missing prices and categories
- ✅ Extract new fields (Manufacturing Date and Country Code) from the Product ID
- ✅ Produce a final, analysis-ready dataset

---

## 📁 Repository Structure

```text
📦 Excel_Product_Sales_Data_Cleaning
 ┣ 📄 Product_Sales_Data_Cleaning_Imputation.xlsx	# Main Workbook (all stages)
 ┗ 📄 Product_Sales_Raw_Dataset.xlsx		   	# Raw Dataset          
 ┗ 📄 Product_Sales_Data_Cleaning_Report.pdf 	    	# Project Documentation
 ┗ 📄 README.md 					# Markdown Documentation
```

---

## 🗂️ Workbook Sheets

| 📄 Sheet | 📝 Purpose | 📏 Size |
|---|---|---|
| 🟥 `Raw_Dataset` | Original uncleaned data | 34 records, 6 columns |
| 🟨 `Data_Cleaning_&_Imputation` | Working sheet with helper columns, formulas and imputation | 31 records, 11 columns |
| 🟩 `Final_Cleaned_Data` | Final cleaned, analysis-ready output | 31 records, 9 columns |

---

## 🧾 Data Dictionary
---

### 🟥 Raw Dataset

| 🏷️ Column | 📖 Description | 🔤 Example |
|---|---|---|
| `Product ID` | Composite key: `DD-MON-CC` (day, month, country code) | `28-JAN-US` |
| `Product Name` | Name of the product | `laptop` |
| `Brand Name` | Manufacturer brand | `Dell` |
| `Price ($)` | Unit price in USD | `1000` |
| `Quantity` | Units sold | `30` |
| `Category` | Product category | `Electronics` |

### 🟩 Final Cleaned Dataset

| 🏷️ Column | 📖 Description | 🔤 Example |
|---|---|---|
| `Product ID` | Original composite key | `28-JAN-US` |
| `Manufacturing Date` | Date extracted from Product ID | `2026-01-28` |
| `Country Code` | Country extracted from Product ID | `US` |
| `Product Name` | Standardised product name | `Laptop` |
| `Brand Name` | Manufacturer brand | `Dell` |
| `Product_Brand` | Product name and brand combined | `Laptop Dell` |
| `Price ($)` | Cleaned price with missing values imputed | `1000` |
| `Quantity` | Units sold | `30` |
| `Category` | Cleaned and completed category | `Electronics` |

---

## 🚨 Data Quality Issues Found
---

| # | ⚠️ Issue | 🔎 Example in Raw Data | 📊 Count |
|---|---|---|---|
| 1 | 🔁 Duplicate records | `21-AUG-CA` Laptop Bag, `16-APR-ES` Bose Headphones, `17-JUN-IN` HP Laptop | 3 |
| 2 | ❓ Missing prices | Sony Headphones, Coleman Camping Tent, Ray-Ban Sunglasses | 3 |
| 3 | ❓ Missing categories | Backpack, Adidas Sneakers, Nespresso Coffee Maker, Xiaomi Fitness Tracker | 4 |
| 4 | ✂️ Truncated category text | `Electroni` instead of `Electronics` | Multiple |
| 5 | 🔠 Inconsistent casing | `laptop`, `smartphone`, `headphones`, `T-shirt` | Multiple |
| 6 | 🧩 Composite Product ID | Date and country code packed into one field | All rows |

---

## 🛠️ Cleaning & Imputation Steps
---

### 1️⃣ 🔁 Remove Duplicates
Three exact duplicate rows were removed, reducing the dataset from **34 to 31 records**.

### 2️⃣ 🔠 Standardise Text
Product names were converted to consistent proper case (`laptop` → `Laptop`, `T-shirt` → `T-Shirt`).

### 3️⃣ ✂️ Split the Product ID
The `Product ID` was split into two fields:
- 📅 **Date** (first 6 characters), converted to a real Excel date → `Manufacturing Date`
- 🌍 **Country Code** (last 2 characters)

### 4️⃣ 🔗 Create `Product_Brand`
Combined Product Name and Brand Name into one descriptive field (e.g. `Coffee Maker Keurig`).

### 5️⃣ 💲 Impute Missing Prices
Missing prices are filled in a two-level fallback:

| Level | 🧠 Logic | 🔤 Example |
|---|---|---|
| 🥇 Level 1 | Average price of other rows with the **same Product Name** | Sony Headphones → average of Headphones → **$250** |
| 🥈 Level 2 | If no match exists, use the **Outdoor category average** | Camping Tent → average of Backpack and Hiking Boots → **$100** |

| 🛍️ Product | ❌ Raw Price | ✅ Imputed Price |
|---|---|---|
| Headphones Sony | blank | 💲 250 |
| Camping Tent Coleman | blank | 💲 100 |
| Sunglasses Ray-Ban | blank | 💲 150 |

### 6️⃣ 🏷️ Fix and Fill Categories
- ✂️ `Electroni` → `Electronics`
- ❓ Blank categories filled by matching products to their category:
  - 🎒 Backpack → `Outdoor` (via the `BackPack_Category` lookup sheet)
  - 👟 Sneakers → `Fashion`
  - ☕ Coffee Maker → `Kitchen`
  - ⌚ Fitness Tracker → `Electronics`

### 7️⃣ ✅ Create the Final Dataset
Helper columns (`Date`, `Price ($)` raw, `Cleaned_Price ($)`) were consolidated, and values were pasted into `Final_Cleaned_Data`.

---

## 🧮 Excel Formulas Used
---

| 🧰 Function | 🎯 Purpose |
|---|---|
| `LEFT()` | Extract the date part from Product ID |
| `RIGHT()` | Extract the country code from Product ID |
| `DATEVALUE()` | Convert text date to an Excel date serial |
| `TEXTJOIN()` | Combine Product Name and Brand Name |
| `IF()` + `ISBLANK()` | Detect missing prices |
| `AVERAGEIF()` | Calculate average price by product or category |
| `IFERROR()` | Fallback when no matching product exists |
| 📋 **Excel Table** (`Product_Sales_DB`) | Structured references for readable formulas |

**📌 Imputation formula (simplified):**

```excel
=IF(ISBLANK([@[Price ($)]]),
    IFERROR(
        AVERAGEIF([Product Name], [@[Product Name]], [Price ($)]),
        AVERAGEIF([Cleaned_Category], "Outdoor", [Price ($)])
    ),
    [@[Price ($)]]
)
```

**📌 Product ID split formulas:**

```excel
=LEFT([@[Product ID]], 6)        → Date          (e.g. 28-JAN)
=DATEVALUE([@Date])              → Manufacturing Date
=RIGHT([@[Product ID]], 2)       → Country Code  (e.g. US)
```

---

## 🔄 Before vs After
---

| 🔍 Aspect | 🟥 Before | 🟩 After |
|---|---|---|
| 📏 Records | 34 | 31 |
| 🔁 Duplicates | 3 | 0 |
| 💲 Missing prices | 3 | 0 |
| 🏷️ Missing or truncated categories | Several | 0 |
| 🔠 Text casing | Inconsistent | Standardised |
| 🧩 Columns | 6 | 9 |
| 📅 Date field | Embedded in ID | Proper date column |
| 🌍 Country field | Embedded in ID | Separate column |

---

## 📈 Final Dataset Summary
---

| 📊 Metric | 🔢 Value |
|---|---|
| 🧾 Total records | 31 |
| 🌍 Countries covered | 12 (US, UK, IN, AU, DE, CA, ES, CN, RU, IT, BR, FR) |
| 💲 Price range | $30 – $1,000 |

**🏷️ Records by category**

| Category | 📦 Records |
|---|---|
| 💻 Electronics | 13 |
| 👗 Fashion | 7 |
| 🍳 Kitchen | 6 |
| 🏕️ Outdoor | 3 |
| 👜 Accessories | 2 |

---

## 🚀 How to Use This Project
---

1. 📥 **Clone or download** this repository
   ```bash
   git clone https://github.com/ManikandanSuriyakumar/Excel_Product_Sales_Data_Cleaning.git
   ```
2. 📂 Open `Product_Sales_Data_Cleaning_Imputation.xlsx` in **Microsoft Excel**
   (Excel 2019 or Microsoft 365 is recommended because `TEXTJOIN` is used)
3. 🔍 Explore the sheets in order: `Raw_Dataset` → `Data_Cleaning_&_Imputation` → `Final_Cleaned_Data`
4. 📤 Use `Final_Cleaned_Data` for analysis, pivot tables, dashboards or Power BI / Python import

---

## 💡 Key Learnings
---

- 🧠 Hierarchical imputation (product-level first, category-level as fallback) keeps estimates realistic
- 🔗 Composite keys such as Product ID can be split into useful analytical fields
- 📋 Excel Tables with structured references make formulas easier to read and maintain
- ✅ Keeping raw, working and final sheets separate preserves an auditable cleaning trail

---

## 🔮 Future Improvements
---

- 🐍 Reproduce the cleaning pipeline in **Python (pandas)**
- 🗄️ Load the final dataset into **SQL** for querying
- 📊 Build an interactive **Power BI / Tableau** dashboard
- 📅 Add the year explicitly to dates (the current `DATEVALUE` call assumes the current year)
- 🤖 Add automated data validation checks

---

## 👤 Author
---

**Manikandan S**

Aspiring Data Analyst | Learning Excel and developing practical data analysis projects.

📧 manisuriyakumar042@gmail.com
💻 https://github.com/ManikandanSuriyakumar

---

## 📚 Acknowledgement
---

This project is developed for educational and learning purposes to demonstrates the Excel data cleaning, transformation and formula-based analysis skills.

---

⭐ If you found this project helpful, please give it a star!