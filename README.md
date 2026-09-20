## Setup

Requires Python 3.11 or newer

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## How to run

Run the steps in this order. Each one needs the output of the previous one.

**Part 1 - explore.** Open `explore.ipynb` in VS Code or Jupyter, select the `.venv` kernel, and choose *Run All*.
Clear the notebook outputs before committing, because they contain rows from the data.

**Part 2 - cleanse.**

```bash
python clean_data.py
```

This reads `moj_sales_extract.xlsx` and writes `cleansed_data.csv` and `cleaning.log`.


**Part 3 - report.**

```bash
python generate_report.py --period 2020Q1
```

This reads `cleansed_data.csv` and writes `output/sales_report_2020Q1.pptx`. Use any quarter in the data, written as
`YYYYQn` (Gregorian calendar quarters). If the quarter is not in the data, the script lists the quarters that are.

The report has seven slides: title, a summary of the quarter, total sales value by district (chart and table),
the same by property category, and a short guide to reading the figures. Every table shows the number of deals behind
each figure, and the change since the previous quarter.

**Part 4 - SQL.** `queries.sql` is meant to be read, not run. It is written for SQL Server against the schema in
`schema.md`.