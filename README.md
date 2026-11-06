# 📊 Data Cleaning and Transformation using Excel – Employee Dataset

## 📌 Business Scenario: Employee Records Management

This project focuses on cleaning and transforming an employee dataset using Microsoft Excel. It demonstrates practical data-cleaning techniques, including handling missing values, removing duplicate records, standardising text, correcting inconsistent data, and transforming employee IDs.

## 📁 Project Overview

**Project Title:** Employee Data Cleaning and Transformation using Excel

### Dataset Description

| Column | Description |
|---|---|
| Employee ID | Code in the format DD-MON-CC |
| Employee Name | Employee's full name |
| Job Title | Employee designation |
| Monthly Salary ($) | Employee's monthly salary |
| Experience (Years) | Total years of experience |
| Department | Employee's department |

## 🔍 Data Quality Issues Found

| Issue | Example | Solution |
|---|---|---|
| Missing Salary | Data Analyst, HR Executive, Accountant | Fill with median |
| Missing Department | Software Developer, Accountant, Designer, Recruiter | Fill based on job title |
| Truncated Text | Financ instead of Finance | Find & Replace |
| Inconsistent Case | manager, MANAGER, analyst | Standardise using PROPER |
| Duplicate Rows | Repeated employee records | Remove Duplicates |
| Combined Values | Employee ID contains date and country code | Split into separate columns |

## 🛠️ Step-by-Step Data Cleaning Process

### 1. Make a Backup

Create three worksheets in Excel:

- `Raw_Dataset` – Original employee dataset.
- `Working_Dataset` – Perform all cleaning operations.
- `Cleaned_Dataset` – Store the final cleaned dataset.

### 2. Remove Duplicate Records

**Excel Navigation:** Data → Remove Duplicates

- Select all columns.
- Click OK to remove duplicate records.

**Result:**
- Original records: 34
- Duplicate records removed: 3
- Final records: 31

### 3. Fix Text Formatting

Use the following formula in a helper column:

```excel
=PROPER(TRIM(B2))
```

**Examples:**

| Original Value | Cleaned Value |
|---|---|
| john smith | John Smith |
| MARY JONES | Mary Jones |
| david kumar | David Kumar |
| anita roy | Anita Roy |

**Purpose:**
- TRIM removes unnecessary spaces.
- PROPER converts text into proper case.

### 4. Fix Truncated Department Names

Use **Find & Replace**.

**Shortcut:** `Ctrl + H`

- Find: `Financ`
- Replace with: `Finance`
- Enable Match entire cell contents.

This corrects the truncated department name without changing existing correct values.

### 5. Handle Missing Department Values

Fill missing department values based on the employee's job title.

| Job Title | Department |
|---|---|
| Software Developer | IT |
| Accountant | Finance |
| Designer | Marketing |
| Recruiter | Human Resources |

### 6. Handle Missing Salary Values

Use median imputation to fill missing salary values.

**Excel Formula:**

```excel
=MEDIAN(D2:D32)
```

**Why Median?**

Employee salaries can have a skewed distribution because of differences in job levels and seniority.

- Median is less affected by extremely high salaries.
- Mean can be influenced by outliers.
- Median provides a more representative central value for skewed salary data.

For better accuracy, calculate the median separately for each department.

### 7. Split Employee ID (Data Transformation)

**Example Employee ID:** `15-JAN-IN`

Use LEFT and RIGHT functions to extract the required information.

| New Column | Excel Formula | Result |
|---|---|---|
| Join Day-Month | `=LEFT(A2,6)` | 15-JAN |
| Country Code | `=RIGHT(A2,2)` | IN |

**Purpose:**
- LEFT extracts the first six characters.
- RIGHT extracts the last two characters.

### 8. Final Data Validation

Perform the following checks:

- No missing salary or department values.
- No duplicate employee records.
- Consistent employee names and department names.
- Salary and experience stored as numeric values.
- Employee ID transformation completed correctly.
- Validate the final dataset using filters and Pivot Tables.

## 📚 Excel Functions and Features Used

| Excel Function / Feature | Purpose |
|---|---|
| Remove Duplicates | Remove repeated employee records |
| Find & Replace | Correct truncated department names |
| TRIM | Remove unnecessary spaces |
| PROPER | Standardise text formatting |
| LEFT | Extract employee ID information |
| RIGHT | Extract country code |
| MEDIAN | Handle missing salary values |
| Go To Special | Identify blank cells |
| Filters | Review and validate records |
| Pivot Tables | Summarise employee information |

## 📈 Project Results

The final cleaned employee dataset contains:

- 31 unique employee records.
- No missing salary or department values.
- Standardised employee names and department names.
- Corrected department information.
- Two additional transformed columns: Join Day-Month and Country Code.

 


