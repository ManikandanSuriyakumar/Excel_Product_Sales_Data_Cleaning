# 🧹📊 Excel - Product Sales Data Cleaning & Imputation

> An Excel-based Data Cleaning project that transforms a messy product
> sales dataset into a clean, analysis-ready table using Excel formulas,
> imputation, and standardisation.

------------------------------------------------------------------------

## 📑 Table of Contents

-   [🎯 Project Overview](#-project-overview)
-   [📁 Repository Structure](#-repository-structure)
-   [🗂️ Workbook Sheets](#️-workbook-sheets)
-   [🧾 Data Dictionary](#-data-dictionary)
-   [🚨 Data Quality Issues Found](#-data-quality-issues-found)
-   [🛠️ Cleaning & Imputation Steps](#️-cleaning--imputation-steps)
-   [🧮 Excel Formulas Used](#-excel-formulas-used)
-   [🔄 Before vs After](#-before-vs-after)
-   [📈 Final Dataset Summary](#-final-dataset-summary)
-   [🚀 How to Use](#-how-to-use-this-project)
-   [💡 Key Learnings](#-key-learnings)
-   [🔮 Future Improvements](#-future-improvements)
-   [👤 Author](#-author)
-   [📚 Acknowledgement](#-acknowledgement)

------------------------------------------------------------------------

## 🎯 Project Overview

This project focuses on cleaning and preparing a **Product Sales
Dataset** covering electronics, fashion, kitchen, outdoor, and accessory
products sold across multiple countries.

The raw dataset contains missing values, inconsistent text, truncated
categories, and duplicate records. This project demonstrates the
complete data cleaning process, from raw data preparation to the final
cleaned dataset using Microsoft Excel.

**🎯 Project Goals**

-   ✅ Remove duplicate records.
-   ✅ Fix inconsistent text casing and truncated values.
-   ✅ Impute missing prices and categories.
-   ✅ Extract new fields (Manufacturing Date and Country Code) from
    Product ID.
-   ✅ Combine product and brand names into a new column.
-   ✅ Prepare a clean, structured, and analysis-ready dataset.

------------------------------------------------------------------------

## 📁 Repository Structure

``` text
📦 Excel_Product_Sales_Data_Cleaning
 ┣ 📄 Product_Sales_Data_Cleaning_Imputation.xlsx
 ┣ 📄 Product_Sales_Raw_Dataset.xlsx
 ┣ 📄 Product_Sales_Data_Cleaning_Report.pdf
 ┗ 📄 README.md
```

------------------------------------------------------------------------

## 🗂️ Workbook Sheets

  ------------------------------------------------------------------------------
  📄 Sheet                       📝 Purpose              📏 Size
  ------------------------------ ----------------------- -----------------------
  🟥 `Raw_Dataset`               Original uncleaned data 34 records, 6 columns

  🟨                             Working sheet with      31 records, 11 columns
  `Data_Cleaning_&_Imputation`   helper columns,         
                                 formulas, and           
                                 imputation              

  🟩 `Final_Cleaned_Data`        Final cleaned,          31 records, 9 columns
                                 analysis-ready output   
  ------------------------------------------------------------------------------

------------------------------------------------------------------------

## 🧾 Data Dictionary

### 🟥 Raw Dataset

  -----------------------------------------------------------------------
  🏷️ Column               📖 Description          🔤 Example
  ----------------------- ----------------------- -----------------------
  `Product ID`            Composite key           `28-JAN-US`
                          containing day, month,  
                          and country code        
                          (`DD-MON-CC`)           

  `Product Name`          Name of the product     `laptop`

  `Brand Name`            Manufacturer brand      `Dell`

  `Price ($)`             Unit price in USD       `1000`

  `Quantity`              Units sold              `30`

  `Category`              Product category        `Electronics`
  -----------------------------------------------------------------------

### 🟩 Final Cleaned Dataset

  🏷️ Column              📖 Description                              🔤 Example
  ---------------------- ------------------------------------------- ---------------
  `Product ID`           Original composite key                      `28-JAN-US`
  `Manufacturing Date`   Date extracted from Product ID              `2026-01-28`
  `Country Code`         Country extracted from Product ID           `US`
  `Product Name`         Standardised product name                   `Laptop`
  `Brand Name`           Manufacturer brand                          `Dell`
  `Product_Brand`        Product name and brand combined             `Laptop Dell`
  `Price ($)`            Cleaned price with missing values imputed   `1000`
  `Quantity`             Units sold                                  `30`
  `Category`             Cleaned and completed product category      `Electronics`

------------------------------------------------------------------------

## 🚨 Data Quality Issues Found

  -----------------------------------------------------------------------
  \#                ⚠️ Issue          🔎 Example in Raw 📊 Count
                                      Data              
  ----------------- ----------------- ----------------- -----------------
  1                 🔁 Duplicate      `21-AUG-CA`       3
                    records           Laptop Bag,       
                                      `16-APR-ES` Bose  
                                      Headphones,       
                                      `17-JUN-IN` HP    
                                      Laptop            

  2                 ❓ Missing prices Sony Headphones,  3
                                      Coleman Camping   
                                      Tent, Ray-Ban     
                                      Sunglasses        

  3                 ❓ Missing        Backpack, Adidas  4
                    categories        Sneakers,         
                                      Nespresso Coffee  
                                      Maker, Xiaomi     
                                      Fitness Tracker   

  4                 ✂️ Truncated      `Electroni`       Multiple
                    category text     instead of        
                                      `Electronics`     

  5                 🔠 Inconsistent   `laptop`,         Multiple
                    casing            `smartphone`,     
                                      `headphones`,     
                                      `T-shirt`         

  6                 🧩 Composite      Date and country  All rows
                    Product ID        code combined     
                                      into one field    
  -----------------------------------------------------------------------

------------------------------------------------------------------------

## 🛠️ Cleaning & Imputation Steps

### 1️⃣ 🔁 Remove Duplicates

Three duplicate records were removed, reducing the dataset from **34 to
31 records**.

### 2️⃣ 🔠 Standardise Text

Product names were converted into consistent proper case.

Examples:

-   `laptop` → `Laptop`
-   `smartphone` → `Smartphone`
-   `T-shirt` → `T-Shirt`

### 3️⃣ ✂️ Split the Product ID

The `Product ID` column was split into two separate fields:

-   📅 **Manufacturing Date:** Extracted from the first six characters
    and converted into an Excel date.
-   🌍 **Country Code:** Extracted from the last two characters.

Example: `28-JAN-US`

### 4️⃣ 🔗 Create `Product_Brand`

Product Name and Brand Name were combined into a single descriptive
column.

Example: `Coffee Maker Keurig`

### 5️⃣ 💲 Impute Missing Prices

Missing prices were filled using a two-level fallback approach.

  -----------------------------------------------------------------------
  Level                   🧠 Logic                🔤 Example
  ----------------------- ----------------------- -----------------------
  🥇 Level 1              Calculate the average   Sony Headphones →
                          price of matching       Average price of
                          product names           Headphones = **\$250**

  🥈 Level 2              If no matching product  Coleman Camping Tent →
                          exists, use the Outdoor **\$100**
                          category average        
  -----------------------------------------------------------------------

**Price Imputation Results**

  🛍️ Product             ❌ Raw Price   ✅ Imputed Price
  ---------------------- -------------- ------------------
  Headphones Sony        Blank          \$250
  Camping Tent Coleman   Blank          \$100
  Sunglasses Ray-Ban     Blank          \$150

### 6️⃣ 🏷️ Fix and Fill Categories

-   ✂️ `Electroni` → `Electronics`
-   ❓ Missing categories were filled using product-based category
    matching.

  🛍️ Product               🏷️ Imputed Category   🔎 Method
  ------------------------ --------------------- ----------------------------------
  Backpack                 Outdoor               `BackPack_Category` lookup sheet
  Adidas Sneakers          Fashion               Product-based matching
  Nespresso Coffee Maker   Kitchen               Product-based matching
  Xiaomi Fitness Tracker   Electronics           Product-based matching

### 7️⃣ ✅ Create the Final Dataset

Helper columns such as `Date`, raw `Price ($)`, and `Cleaned_Price ($)`
were consolidated into the final dataset.

The cleaned values were transferred to the `Final_Cleaned_Data` sheet,
resulting in a structured, analysis-ready table.

------------------------------------------------------------------------

## 🧮 Excel Formulas Used

  -----------------------------------------------------------------------
  🧰 Function                         🎯 Purpose
  ----------------------------------- -----------------------------------
  `LEFT()`                            Extract the date portion from
                                      Product ID

  `RIGHT()`                           Extract the country code from
                                      Product ID

  `DATEVALUE()`                       Convert text date into an Excel
                                      date serial

  `TEXTJOIN()`                        Combine Product Name and Brand Name

  `IF()` + `ISBLANK()`                Identify missing prices

  `AVERAGEIF()`                       Calculate average price based on
                                      product or category

  `IFERROR()`                         Apply a fallback when no matching
                                      product is found

  📋 **Excel Table**                  Use structured references for
  (`Product_Sales_DB`)                readable and maintainable formulas
  -----------------------------------------------------------------------

### 📌 Price Imputation Formula (Simplified)

``` excel
=IF(ISBLANK([@[Price ($)]]),
   IFERROR(
       AVERAGEIF([Product Name], [@[Product Name]], [Price ($)]),
       AVERAGEIF([Cleaned_Category], "Outdoor", [Price ($)])
   ),
   [@[Price ($)]]
)
```

### 📌 Product ID Extraction Formulas

``` excel
=LEFT([@[Product ID]], 6)        → Extract Date
=DATEVALUE([@Date])              → Convert to Manufacturing Date
=RIGHT([@[Product ID]], 2)       → Extract Country Code
```

------------------------------------------------------------------------

## 🔄 Before vs After

  -----------------------------------------------------------------------
  🔍 Aspect               🟥 Before               🟩 After
  ----------------------- ----------------------- -----------------------
  📏 Total Records        34                      31

  🔁 Duplicate Records    3                       0

  💲 Missing Prices       3                       0

  🏷️ Missing or Truncated Several                 0
  Categories                                      

  🔠 Text Casing          Inconsistent            Standardised

  🧩 Total Columns        6                       9

  📅 Date Field           Embedded in Product ID  Separate Date Column

  🌍 Country Field        Embedded in Product ID  Separate Country Code
                                                  Column
  -----------------------------------------------------------------------

------------------------------------------------------------------------

## 📈 Final Dataset Summary

  📊 Metric              🔢 Value
  ---------------------- ----------
  🧾 Total Records       31
  🌍 Countries Covered   12
  💲 Minimum Price       \$30
  💲 Maximum Price       \$1,000

**🌍 Countries Covered**

US, UK, IN, AU, DE, CA, ES, CN, RU, IT, BR, FR.

### 🏷️ Records by Category

  Category           📦 Records
  ---------------- ------------
  💻 Electronics             13
  👗 Fashion                  7
  🍳 Kitchen                  6
  🏕️ Outdoor                  3
  👜 Accessories              2
  **Total**              **31**

------------------------------------------------------------------------

## 🚀 How to Use This Project

1.  📥 **Clone or download** this repository.

    ``` bash
    git clone https://github.com/ManikandanSuriyakumar/Excel_Product_Sales_Data_Cleaning.git
    ```

2.  📂 Open `Product_Sales_Data_Cleaning_Imputation.xlsx` in **Microsoft
    Excel**.

    *(Excel 2019 or Microsoft 365 is recommended because the workbook
    uses `TEXTJOIN()`.)*

3.  🔍 Explore the workbook sheets in the following order:

    `Raw_Dataset` → `Data_Cleaning_&_Imputation` → `Final_Cleaned_Data`

4.  📤 Use the `Final_Cleaned_Data` sheet for further analysis, Pivot
    Tables, dashboards, or importing into Power BI and Python.

------------------------------------------------------------------------

## 💡 Key Learnings

-   🧠 Applied hierarchical imputation using product-level averages and
    category-level fallback methods.
-   🔗 Learned to extract meaningful information from composite
    identifiers.
-   📋 Practised Excel Tables and structured references to simplify
    formulas.
-   🧹 Identified and corrected missing values, duplicate records,
    inconsistent text, and incorrect categories.
-   ✅ Understood the importance of maintaining separate raw, working,
    and final datasets for a traceable cleaning process.

------------------------------------------------------------------------

## 🔮 Future Improvements

-   🐍 Reproduce the data cleaning process using **Python (Pandas)**.
-   🗄️ Load the final cleaned dataset into **SQL** for querying and
    analysis.
-   📊 Develop an interactive **Power BI / Tableau** dashboard.
-   📅 Include the year explicitly in the date extraction process.
-   🤖 Implement automated data validation and quality checks.

------------------------------------------------------------------------

## 👤 Author

**Manikandan S**

Aspiring Data Analyst \| Learning Excel and developing practical data
analysis projects.

📧 Email: manisuriyakumar042@gmail.com

💻 GitHub:
[ManikandanSuriyakumar](https://github.com/ManikandanSuriyakumar)

------------------------------------------------------------------------

## 📚 Acknowledgement

This project is developed for educational and learning purposes to
demonstrate Excel-based data cleaning, transformation, imputation, and
formula-based data preparation techniques.

------------------------------------------------------------------------

⭐ If you found this project helpful, please give it a star!
