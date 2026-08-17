# Headcount Splitter

A basic Streamlit app that:

1. Uploads an `.xlsx` headcount workbook.
2. Lets you select the worksheet and header row.
3. Lets you choose the department column.
4. Shows department row counts.
5. Creates one `.xlsx` file per selected department.
6. Downloads all generated files in a ZIP.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL shown by Streamlit.

## Notes

- Version 1 supports `.xlsx` only.
- It copies the selected sheet's headers, matching data rows, common cell formatting and column widths.
- It does not yet preserve advanced Excel objects such as charts, slicers, macros, external links or complex merged layouts.
