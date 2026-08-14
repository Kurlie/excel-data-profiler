````markdown
# Excel Data Profiler

A lightweight Excel data profiler for quickly identifying completeness, consistency, patterns, anomalies, and potential data quality issues.

Built for analysts, consultants, small businesses, and anyone who needs to understand an Excel dataset without setting up a database, writing SQL, using Python, or deploying an enterprise data quality platform.

---

## What It Does

Excel Data Profiler analyzes the active worksheet in your Excel workbook and creates a structural and statistical profile of the dataset.

It can help surface things such as:

- Missing or blank values
- Completeness by column
- Distinct and duplicate values
- Exact duplicate records
- Mixed data types
- Inferred data types and confidence
- Minimum and maximum values
- Text length variation
- Leading or trailing whitespace
- Most common values
- Statistical outliers
- Email pattern inconsistencies
- Phone number pattern inconsistencies
- IPv4 pattern inconsistencies
- Potential SSN patterns
- Potential sensitive-data fields
- Potential identifier fields
- Future or unusual date values
- Ambiguous date-like text
- Excel formula errors
- Potential Excel conversion risks
- Long numeric identifiers that may exceed Excel's numeric precision
- Scientific-notation risks in identifier-like fields

> **The goal is not to tell you whether your data is correct.**
>
> The goal is to quickly show you what is in your data and what may deserve further investigation.

---

## Who Is This For?

This tool is intentionally designed for smaller, everyday datasets that already live in Excel.

Examples include:

- Customer lists
- Product files
- Employee exports
- CRM exports
- Accounting extracts
- Vendor files
- Inventory data
- Marketing lists
- Operational spreadsheets
- CSV files imported into Excel
- Data received from customers, vendors, or business partners

It is **not intended to replace enterprise-scale data profiling or data quality platforms**.

---

# Getting Started

## Requirements

You will need:

- Microsoft Excel with **Office Scripts** support
- Access to the **Automate** tab in Excel
- A Microsoft 365 account that allows Office Scripts
- OneDrive or SharePoint access for Office Script storage

If you do not see the **Automate** tab in Excel, Office Scripts may not be available for your Microsoft 365 account or may be disabled by your organization.

---

# Installing the Profiler

The profiler is distributed as an Office Script file:

```text
data-profiler.osts
````

## Option 1: Add the `.osts` File to Your Office Scripts Folder

### Step 1: Download the Script

Download:

```text
data-profiler.osts
```

from this GitHub repository.

You can do this by opening the file in GitHub and downloading the raw file, or by downloading/cloning the repository.

---

### Step 2: Open OneDrive

Open your Microsoft OneDrive and navigate to:

```text
Documents
└── Office Scripts
```

Office Scripts are normally stored in the `Documents/Office Scripts` folder.

> If you do not already have an `Office Scripts` folder, open Excel and create or save an Office Script from the **Automate** tab first. Excel should create the appropriate storage location.

---

### Step 3: Add the Script

Place:

```text
data-profiler.osts
```

inside:

```text
Documents/Office Scripts
```

in OneDrive.

---

### Step 4: Open Excel

Open or restart Excel.

Go to:

**Automate → View Scripts**

The Excel Data Profiler should now appear in your available Office Scripts.

---

# Running the Profiler

## Step 1: Open Your Dataset

Open the Excel workbook containing the data you want to profile.

The profiler analyzes the **active worksheet**.

Make sure you are viewing the worksheet containing your source dataset before running the script.

### Expected Dataset Structure

Your data should look roughly like this:

| Customer_ID | Name        | Email                                       | State | Order_Amount |
| ----------- | ----------- | ------------------------------------------- | ----- | -----------: |
| CUST001     | Jane Smith  | [jane@example.com](mailto:jane@example.com) | OH    |       125.50 |
| CUST002     | John Davis  | [john@example.com](mailto:john@example.com) | MI    |        74.99 |
| CUST003     | Sarah Jones |                                             | OH    |       210.00 |

The profiler assumes the **first row of the used range contains column headers**.

The following can all vary dynamically:

* Number of rows
* Number of columns
* Column names
* Column order
* Data types

---

## Step 2: Check Your Worksheet Names

> [!IMPORTANT]
> Before running the profiler, make sure your workbook does **not** already contain worksheets named:
>
> * `Profile Summary`
> * `Column Profile`
> * `Data Quality Issues`

The profiler recreates these worksheets each time it runs.

**Any existing worksheets with these exact names will be deleted and replaced.**

---

## Step 3: Select Your Source Worksheet

Click the worksheet containing the dataset you want to analyze.

Do **not** select one of the profiler output worksheets.

---

## Step 4: Run the Script

In Excel:

1. Select **Automate**
2. Select **View Scripts**
3. Find **Excel Data Profiler**
4. Select the script
5. Select **Run**

The profiler will analyze the active worksheet and generate the profiling results.

Processing time will vary depending on the size of your dataset.

---

# Understanding the Results

The profiler creates three worksheets:

1. `Profile Summary`
2. `Column Profile`
3. `Data Quality Issues`

---

## 1. Profile Summary

Provides a high-level overview of the dataset, including:

* Source worksheet
* Profile date/time
* Total records
* Total columns
* Total populated cells
* Total blank/null cells
* Overall completeness
* Empty columns
* Duplicate records
* Duplicate record percentage
* Columns requiring review
* Mixed data type issues
* Pattern inconsistencies
* Excel error issues
* Whitespace issues
* Potential identifier columns
* Potential sensitive-data columns
* Excel conversion risks

The summary also highlights some of the most significant issues found during profiling.

**This is the best place to start when reviewing an unfamiliar dataset.**

---

## 2. Column Profile

Provides detailed statistics for every column in the dataset.

Depending on the data, this can include:

* Record count
* Populated count
* Blank count
* Completeness %
* Distinct count
* Uniqueness %
* Duplicate value count
* Detected data type
* Data type confidence
* Observed data types
* Minimum value
* Maximum value
* Minimum text length
* Maximum text length
* Average text length
* Most common value
* Most common value frequency
* Example values
* Potential identifier classification
* Potential sensitive-data classification
* Mixed data type indicators
* Whitespace exceptions
* Pattern type
* Pattern conformance
* Pattern exceptions
* Potential Excel conversion risks
* Overall quality flag
* Profiling notes

Use this worksheet when you want to understand the characteristics of individual fields.

---

## 3. Data Quality Issues

Provides a consolidated list of items that may deserve further investigation.

Each issue includes:

* Severity
* Column
* Issue type
* Description
* Records affected
* Percent of applicable values
* Recommended review

Example observations might include:

```text
Email contains 2 populated values that do not match the detected Email pattern.
```

```text
Customer_ID contains values with leading or trailing whitespace.
```

```text
Order_Amount contains potential statistical outliers.
```

```text
Account_Number contains long numeric identifier values that may be affected by Excel numeric precision.
```

These are **profiling observations**, not declarations that the values are incorrect.

---

# Your Source Data Is Not Modified

The profiler is designed to treat the active source worksheet as **read-only**.

It reads the dataset into memory and performs profiling against those in-memory values.

The profiler does not intentionally:

* Delete source records
* Remove duplicates
* Trim source values
* Replace blanks
* Correct dates
* Change data types
* Modify formulas
* Add source columns
* Sort source data
* Reformat the source worksheet

Only the three profiler output worksheets are written to.

As with any tool operating on important data, keeping an original copy of your workbook is still a good practice.

---

# Important Excel Limitation

This profiler analyzes the data **as Excel currently sees it**.

Excel may automatically interpret or convert certain values when a file is opened or imported.

Examples include:

```text
000123456
03-04
3E10
12345678901234567
```

Depending on how the data was imported, Excel may interpret these as:

* Numbers
* Dates
* Scientific notation
* Other Excel-formatted values

The profiler can identify **potential Excel conversion risks**, but it cannot recover an original source value after Excel has already changed it.

If preserving raw values is critical, review the original:

* CSV
* Text file
* Database extract
* Source system

---

# Dataset Size

This profiler is intentionally designed for **smaller and medium-sized Excel datasets**.

It performs most analysis in memory for better performance, but Excel and Office Scripts still have practical limits.

The script will issue an informational warning when a dataset contains more than:

```text
100,000 records
```

Larger datasets may still run, but processing can take longer.

If you are profiling millions of records or warehouse-scale datasets, a database or dedicated data profiling platform is a better fit.

---

# Privacy and Sensitive Data

The profiler does not call an external API or intentionally send your dataset to an external profiling service.

Profiling is performed through **Excel and Office Scripts**.

The profiler may identify patterns that *could* represent sensitive data, including:

* Email addresses
* Phone numbers
* IP addresses
* SSN-like patterns
* Tax ID-like patterns
* Credit card-like patterns

These classifications are intentionally described as **potential** sensitive data.

Pattern detection alone cannot determine whether a value represents a real person, real account, or actual sensitive information.

The profiler also avoids displaying detected sensitive values directly in issue descriptions.

> Always follow your organization's security, privacy, retention, and data-handling requirements.

---

# What This Tool Does Not Do

Excel Data Profiler is a **profiling tool**, not a business-rule engine.

It does not know what your data *should* contain.

For example, the profiler may tell you:

> **Future date values detected.**

It will not automatically tell you:

> **This date is wrong.**

A future date may be perfectly valid depending on the business process.

Likewise:

* Statistical outliers are not automatically errors
* Duplicate values are not automatically errors
* Missing values are not automatically errors
* Mixed data types are not automatically errors
* Pattern inconsistencies are not automatically errors

They are signals worth investigating.

Business-specific validation requires business context and defined data quality rules.

---

# Quality Flags

Columns may receive one of the following profiling flags:

### Good

No significant structural or statistical concerns were detected.

### Review

One or more observations may deserve investigation.

### Concern

More significant structural or consistency issues were detected.

> A quality flag does **not** determine whether the underlying data is correct for its intended business use.

---

# Why I Built This

Enterprise data quality platforms are powerful, but not every data problem requires enterprise software.

Sometimes you receive an Excel file or CSV and simply need to answer:

* What is actually in this dataset?
* How complete is it?
* Are there duplicates?
* Are the columns consistent?
* What data types are present?
* Are there unusual values?
* Did Excel potentially interpret something incorrectly?
* What should I investigate first?

**Excel Data Profiler is intended to make answering those questions easier.**

---

# Feedback and Contributions

This project is still evolving.

If you find:

* A bug
* A false positive
* A profiling scenario that is not handled well
* A useful profiling capability that should be considered

please open an **Issue** in this GitHub repository.

Feedback from real-world datasets is especially helpful.

---

# License

This project is available under the **MIT License**.

See the `LICENSE` file for details.

---

# Disclaimer

This tool is provided as a lightweight data profiling utility.

Profiling results are based on structural, statistical, and pattern-based analysis and should not be treated as proof that data is correct or incorrect.

Always validate important findings against:

* The intended business use
* The source system
* Applicable business rules
* Applicable data quality requirements

```
```
