# HR Analytics Dashboard in Excel

An interactive **HR Analytics Dashboard built in Microsoft Excel** to transform messy employee data into a clean, analysis-ready dataset and provide insights into workforce composition, attrition, compensation, engagement, performance, and training.

The project demonstrates an end-to-end Excel analytics workflow:

**Raw HR Data → Data Cleaning → Lookup & Master Tables → Analysis-Ready Dataset → PivotTables → Interactive Dashboard**

The final analytical dataset contains **1,000 employee records**.

---

## Project Overview

The source HR dataset contains employee information across multiple European locations and includes deliberate data-quality issues such as:

- Extra spaces
- Mixed text casing
- Misspelled values
- Missing values
- Duplicate employee IDs
- Invalid and mixed date formats
- Inconsistent employment status values
- Text-based performance ratings
- Negative or invalid values
- Missing mapping information

The objective was to clean and standardise the data, enrich it using reference tables, prepare an analysis-ready dataset, and create an interactive HR dashboard for workforce analysis.

---

## Dashboard

> Add your dashboard screenshot here after placing it inside the `screenshots/` directory.

```markdown
![HR Analytics Dashboard](screenshots/hr-analytics-dashboard.png)
```

The dashboard provides a high-level view of workforce performance and allows HR information to be explored across different organisational dimensions.

---

## Key HR KPIs

The dashboard includes key workforce metrics calculated from the cleaned dataset.

| KPI | Result |
|---|---:|
| Total Employees | **1,000** |
| Attrition Rate | **15.3%** |
| Average Salary | **81,886** |
| Average Training Hours | **23.89** |
| Average Employee Engagement Score | **62.67** |

These KPIs provide an overall snapshot of workforce size, employee turnover, compensation, training participation, and engagement.

---

## Dashboard Analysis

The workbook uses PivotTables and dashboard visualisations to analyse several HR dimensions.

### Workforce by Department

Employee headcount is analysed across departments to understand workforce distribution.

Examples from the dataset include:

- Engineering — **164 employees**
- Sales — **121 employees**
- Customer Success — **111 employees**
- Operations — **91 employees**
- Product — **77 employees**

### Attrition by Department

Attrition is compared across departments to identify areas experiencing higher employee turnover.

Examples include:

- Customer Success — **27.0%**
- Corporate Strategy — **25.0%**
- Security — **23.1%**
- Product — **20.8%**
- Operations — **16.5%**

### Salary by Job Level

Average salary is analysed across employee job levels.

Examples include:

- L7 — **234,800**
- L6 — **166,096**
- L5 — **132,114**
- L4 — **101,663**
- L3 — **74,345**

This provides a clear view of compensation progression across organisational levels.

### Employee Engagement

Average employee engagement is analysed across business units, including:

- Corporate Services
- Data & AI
- Operations
- Product & Technology
- Revenue

### Training Analysis

Employees are grouped into training categories to analyse workforce development activity:

- **High Training:** 279 employees
- **Moderate Training:** 601 employees
- **Low Training:** 120 employees

---

## Data Cleaning & Transformation

A major part of the project involved transforming deliberately messy HR data into a consistent analytical dataset.

### Text Standardisation

Employee and organisational fields were cleaned by:

- Removing unnecessary spaces
- Correcting inconsistent casing
- Standardising names and categorical values
- Correcting inconsistent department and location values

Excel functions used include:

```excel
TRIM()
PROPER()
UPPER()
```

### Lookup-Based Standardisation

Dedicated lookup and master tables were created to map inconsistent source values to standard business values.

The workbook includes reference data for:

- Departments
- Managers
- Locations
- Performance ratings
- Cost centres
- Business units

Lookup techniques include:

```excel
XLOOKUP()
VLOOKUP()
IFERROR()
```

These mappings help ensure that inconsistent source values are converted into standard organisational values before analysis.

### Date Cleaning

The raw dataset contains inconsistent and invalid date values.

Cleaning logic was used to convert valid values into consistent Excel dates while identifying or handling invalid entries.

This applies to fields such as:

- Date of Birth
- Joining Date
- Exit Date

### Employment Status Standardisation

Different representations of employment status were converted into consistent categories such as:

- `Active`
- `Exited`

For example, inconsistent values such as `A`, `active`, and `Active` are standardised before analysis.

### Performance Rating Standardisation

Performance ratings appearing as both numeric and text values were mapped to a standard rating system.

Examples include:

| Raw Value | Standard Rating | Rating Label |
|---|---:|---|
| ONE | 1 | Unsatisfactory |
| TWO | 2 | Needs Improvement |
| THREE | 3 | Meets Expectations |
| FOUR | 4 | Exceeds Expectations |
| FIVE | 5 | Outstanding |

The performance master also groups employees into categories such as:

- Low Performer
- Core Performer
- High Performer

### Data Validation & Quality Checks

Cleaning logic was also used to identify data-quality problems such as:

- Duplicate IDs
- Invalid values
- Missing mappings
- Inconsistent categories
- Invalid dates

The cleaned output is then used to create the analysis-ready PivotTable source.

---

## Data Preparation Architecture

The workbook separates raw data, cleaning logic, reference data, analysis data, and presentation rather than performing all operations directly inside the dashboard.

```text
Raw Employee Data
        │
        ▼
Data Cleaning & Standardisation
        │
        ├── Department Lookups
        ├── Manager Mapping
        ├── Location Master
        ├── Performance Master
        └── Cost Centre Mapping
        │
        ▼
Analysis-Ready HR Dataset
        │
        ▼
PivotTables
        │
        ▼
Interactive HR Dashboard
```

This separation makes the workbook easier to maintain, validate, and extend.

---

## Workbook Structure

The Excel workbook contains the following main worksheets:

```text
Assignment_1_HR_Analytics.xlsx
│
├── 01_raw_data
├── 01_clean_data
├── 03_lookups
├── 04_departments
├── 05_manager_mapping
├── 06_location_master
├── 07_performance_master
├── 09_cost_center
├── 10_pivot_source
├── PIVOT
└── Dashboard
```

### `01_raw_data`

Contains the original messy HR extract before cleaning and standardisation.

### `01_clean_data`

Contains the cleaned and transformed employee data.

### Lookup & Master Sheets

Reference tables support standardisation and enrichment:

- `03_lookups`
- `04_departments`
- `05_manager_mapping`
- `06_location_master`
- `07_performance_master`
- `09_cost_center`

For example, the location master converts inconsistent location values into standard locations while deriving the corresponding **country, state, and European region**.

### `10_pivot_source`

Contains the clean, analysis-ready HR dataset used as the source for PivotTables and PivotCharts.

### `PIVOT`

Contains the PivotTables used to calculate KPIs and support dashboard visualisations.

### `Dashboard`

Contains the final interactive HR Analytics Dashboard.

---

## Excel Skills Demonstrated

This project demonstrates practical use of:

### Data Cleaning

- `TRIM`
- `PROPER`
- `UPPER`
- `VALUE`
- `IF`
- `IFERROR`

### Data Integration & Lookups

- `XLOOKUP`
- `VLOOKUP`
- Reference/master tables
- Cross-sheet mappings

### Analysis

- PivotTables
- PivotCharts
- Aggregations
- Employee counts
- Average calculations
- Attrition analysis
- Salary analysis
- Engagement analysis
- Training analysis

### Dashboarding

- KPI cards
- Interactive charts
- Slicers
- Workforce segmentation
- HR performance reporting

---

## Business Questions Addressed

The dashboard helps answer questions such as:

- How large is the current workforce?
- What is the overall employee attrition rate?
- Which departments have higher attrition?
- How is the workforce distributed across departments?
- How does average salary vary by job level?
- How does employee engagement vary across business units?
- How much training are employees receiving?
- How are employees distributed across training categories?
- How can inconsistent HR source data be standardised for reliable reporting?

---

## Tools Used

- **Microsoft Excel**
- Excel Tables
- PivotTables
- PivotCharts
- Slicers
- Lookup functions
- Data-cleaning formulas
- Reference/master tables

---

## Project Structure

```text
excel-hr-analytics-dashboard/
│
├── README.md
│
├── data/
│   └── Assignment_1_HR_Analytics.xlsx
│
└── screenshots/
    └── hr-analytics-dashboard.png
```

---

## What I Learned

This project strengthened my understanding of the complete analytics process in Excel rather than focusing only on dashboard visualisation.

Key areas of learning included:

- Cleaning inconsistent real-world-style HR data
- Designing reusable lookup and master tables
- Standardising categorical data across multiple fields
- Preparing analysis-ready datasets
- Building PivotTables for HR KPIs
- Analysing workforce attrition and compensation
- Creating interactive Excel dashboards
- Separating raw data, transformation logic, analysis, and presentation layers

---

## Future Improvements

Possible future extensions include:

- Rebuilding the dashboard in **Power BI**
- Creating DAX measures for the main HR KPIs
- Automating data preparation using **Power Query**
- Adding time-based attrition and hiring trends
- Expanding demographic and performance analysis
- Connecting the dashboard to a database rather than a static workbook

---

## Author

**Nischal Shakya**

Software Engineer expanding into Data Engineering & Analytics.

- [GitHub](https://github.com/shakya-nischal)
- [LinkedIn](https://www.linkedin.com/in/nischal-shakya-79860a178/)
- [Portfolio](https://nischal-shakya.vercel.app/)
