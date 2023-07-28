# Monitoring-for-Excel-logs
This script compares the latest xlsx file with the penultimate one (file format: 2023-12-31_name.xlsx) and creates an xslx (file format: 2023-12-31_name_log .xlsx) file with information from the last file but with cells highlighted in red that differ in two files.

## How to use
The only thing you need to change is the format or name of your files you can find the responsible code on the 44 line 
```python
all_files = glob.glob("????-??-??_df*.xlsx")
```
