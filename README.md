# 📊 Data Analytics – Excel Assignment 1: Data Exploration

## 📌 About the Project

This project is part of my **Data Analytics learning journey**, focusing on **Microsoft Excel for data exploration and basic analysis**.

The assignment uses a **Product Dataset** containing information about products, including Product ID, Product Name, Brand Name, Quantity, Category, and Price.

The objective of this project is to build foundational Excel skills used by Data Analysts for **data summarization, conditional analysis, logical classification, and text manipulation**.

---

## 🎯 Objectives

Through this assignment, I practiced:

* Exploring and summarizing a dataset
* Performing calculations using Excel functions
* Applying logical conditions to classify data
* Performing conditional calculations
* Extracting information from text fields
* Creating new calculated columns
* Using Excel formulas for basic data analysis

---

## 📂 Dataset

The dataset contains the following attributes:

| Column       | Description                        |
| ------------ | ---------------------------------- |
| Product ID   | Unique identifier for each product |
| Product Name | Name of the product                |
| Brand Name   | Brand associated with the product  |
| Quantity     | Number of products                 |
| Category     | Product category                   |
| Price        | Price of the product               |

---

## 🧮 Tasks Performed

### 1. Basic Data Exploration

Used Excel functions to calculate:

* **Total Price** of all products
* **Number of Products**
* **Average Product Price**

Functions used:

```excel
SUM()
COUNT()
AVERAGE()
```

---

### 2. Minimum and Maximum Price

Identified the lowest and highest product prices using:

```excel
MIN()
MAX()
```

This helped understand the price range within the dataset.

---

### 3. Logical Function – IF

Created a new column called **Price Range** to classify products according to their price.

**Business Rule:**

* Price ≥ $500 → **High Price**
* Price < $500 → **Standard Price**

Formula used:

```excel
=IF(F2>=500,"High Price","Standard Price")
```

This demonstrates the use of logical conditions for data classification.

---

### 4. Conditional Functions – SUMIF and COUNTIF

Performed conditional analysis using:

#### Total Price of Electronics Products

Used `SUMIF()` to calculate the total price of products belonging to the **Electronics** category.

```excel
=SUMIF(E:E,"Electronics",F:F)
```

#### Products Priced Below $100

Used `COUNTIF()` to determine the number of products with a price below $100.

```excel
=COUNTIF(F:F,"<100")
```

These functions helped practice conditional aggregation and filtering concepts.

---

### 5. Text Functions – LEFT, RIGHT and MID

Extracted specific information from the **Product ID** column using Excel text functions.

#### Day

Extracted the first two characters of Product ID using:

```excel
=LEFT(A2,2)
```

#### Country Code

Extracted the last two characters of Product ID using:

```excel
=RIGHT(A2,2)
```

#### Month

Extracted characters 4 to 6 from Product ID using:

```excel
=MID(A2,4,3)
```

These functions demonstrate how Data Analysts can extract useful information from structured text fields.

---

## 🛠️ Excel Skills Demonstrated

* `SUM`
* `COUNT`
* `AVERAGE`
* `MIN`
* `MAX`
* `IF`
* `SUMIF`
* `COUNTIF`
* `LEFT`
* `RIGHT`
* `MID`
* Basic data exploration
* Conditional analysis
* Data classification
* Text extraction
* Formula-based analysis

---

## 📈 Key Learning Outcomes

Through this project, I gained practical experience in using Excel formulas to:

* Summarize product data
* Analyze numerical values
* Apply business rules using logical functions
* Perform category-based calculations
* Extract meaningful information from Product IDs
* Create new analytical columns from existing data

This project strengthened my understanding of **Excel as a foundational tool for data analysis**.

---

## 📁 Project Files

* `Excel Assignment 1 - Data Exploration.xlsx` – Dataset and completed analysis
* `Excel Assignment-1.pdf` – Assignment instructions

---

## 🚀 About My Data Analytics Journey

I am currently developing my skills to pursue a career as a **Data Analyst**.

My learning roadmap includes:

**Excel → SQL → Power BI → Python → Data Analytics Projects**

I am building this GitHub repository to document my learning journey and create a portfolio of practical data analysis projects.

---

## 👩‍💻 Author

**Monisha**

Aspiring Data Analyst | Excel | SQL | Power BI | Python

---

⭐ This repository is part of my ongoing journey to develop practical **Data Analytics and Business Intelligence skills**.
