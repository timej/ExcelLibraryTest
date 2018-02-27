import pandas as pd

def pdread(in_path, out_path):
    df = pd.read_excel(in_path, header=None)
    df.to_csv(out_path, header=False, index=False)

def pdwrite(in_path, out_path):
    df = pd.read_csv(in_path, sep='\t', header=None)
    df.to_excel(out_path, header=False, index=False)

'''
def openpyxl_write(in_path, out_path):
    df = pd.read_csv(in_path, sep='\t', header=None)
    df.to_excel(out_path, header=False, index=False, engine='openpyxl')
'''
def openpyxl_write(in_path, out_path):
    from openpyxl import Workbook
    from openpyxl.utils.dataframe import dataframe_to_rows

    df = pd.read_csv(in_path, sep='\t', header=None)
    
    wb = Workbook(write_only=True)
    ws = wb.create_sheet()

    for r in dataframe_to_rows(df, index=True, header=True):
        ws.append(r)
    wb.save(out_path)


if __name__ == "__main__":
    import sys
    import os
    if not os.path.exists('App_Data/DataOut'):
        os.mkdir('App_Data/DataOut')
    if '-r' in sys.argv:
        if '-x' in sys.argv: 
            pdread('App_Data/DataIn/am0411.xlsx', 'App_Data/DataOut/px.txt')
        else:
            pdread('App_Data/DataIn/am0411.xls', 'App_Data/DataOut/ps.txt')
    else:
        if '-x' in sys.argv:
            pdwrite('App_Data/DataIn/am0411.txt', 'App_Data/DataOut/px.xlsx')
        elif '-o' in sys.argv:
            openpyxl_write('App_Data/DataIn/am0411.txt', 'App_Data/DataOut/po.xlsx')
        else:
            pdwrite('App_Data/DataIn/am0411.txt', 'App_Data/DataOut/ps.xls')