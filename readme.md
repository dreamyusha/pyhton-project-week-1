# Global Superstore — Week 1 Data Analytics & Excel Dashboard

Automated data cleaning, statistical profiling, and interactive Excel dashboard generation for the **Global Superstore (`global_superstore_2016.xlsx`)** dataset.

## Key Findings
- **Total Revenue:** `$12,642,501.91` across `51,290` line items (`25,728` unique orders)
- **Data Cleaning:** `0` exact row duplicates found; `41,296` missing international `Postal Code` entries imputed as `"Not Available"`
- **Sales Descriptive Statistics:** Mean = `$246.49` | Median = `$85.05` | Mode = `$12.96` | Sample Variance = `237,719.98`
- **Top Region:** Western Europe (`$1,731,929.67`)
- **Top Category:** Technology (`$4,744,557.50` — `37.53%` of total revenue)
- **Yearly Trend:** Consistent growth from `$2.26M` (2012) to `$4.30M` (2015)

## How to Run
```bash
pip install -r requirements.txt
python week1_sales_analysis.py