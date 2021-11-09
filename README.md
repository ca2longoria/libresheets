# libresheets
Quick and simple library to parse cell text data from a LibreOffice Calc ods spreadsheet file.

```python
from libresheets import SimpleSheets

sheets = SimpleSheets('spreadsheet.ods')
# For raw table.
sheets.sheets()
# For json-safe raw table.
sheets.clean_sheets()
```

This library is very simple.  Please forgive the utter lack of robust testing.  So I command.
