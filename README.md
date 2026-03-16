### 2. Automated Monthly Report Generator  
**File:** [monthly_report_generator.py](monthly_report_generator.py)

Merges data from multiple Excel or CSV files, calculates useful summaries (totals, averages, counts), and outputs a clean, formatted report workbook.

**Features:**
- Combines files automatically (supports .xlsx and .csv)
- Flexible grouping & aggregation (e.g., sum sales by category)
- Generates summary stats with pandas
- Creates a nicely formatted Excel report using openpyxl:
  - Bold title with current month/year
  - Clean table layout for summary data
  - Auto-formatted headers and data rows
- Fallback to generic stats if no specific columns found
- Error handling for mismatched files/columns

**Perfect for:**
- Monthly sales/inventory reports
- Combining department data
- Quick KPI dashboards from scattered spreadsheets

**Before / After Example**

(Add screenshots here)  
Before: Separate messy monthly files  
After: Single polished report with summary table

**Quick Run Example**

```python
from monthly_report_generator import generate_monthly_report

files = [
    "sales_january.xlsx",
    "sales_february.csv",
    "sales_march.xlsx"
]

generate_monthly_report(files, output_file="monthly_sales_summary.xlsx")
# → Creates monthly_sales_summary.xlsx with merged & summarized data
