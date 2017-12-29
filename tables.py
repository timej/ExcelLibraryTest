import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo

# http://openpyxl.readthedocs.io/en/default/pandas.html#

path = 'App_Data/DataIn/population.txt'
df = pd.read_csv(path, sep='\t')

wb = Workbook(write_only=True)
ws = wb.create_sheet()

for r in dataframe_to_rows(df, index=False, header=True):
    ws.append(r)

# http://openpyxl.readthedocs.io/en/default/worksheet_tables.html

tab = Table(displayName="Table1", ref="A1:F" + str(len(df.index) + 1))

# Add a default style with striped rows and banded columns
style = TableStyleInfo(name="TableStyleLight15", showFirstColumn=False,
                       showLastColumn=False, showRowStripes=True, showColumnStripes=False)
tab.tableStyleInfo = style
ws.add_table(tab)

wb.save('App_Data/DataOut/population.xlsx')

