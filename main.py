import pandas as pd

def pdread(in_path, out_path):
    df = pd.read_excel(in_path, header=None)
    df.to_csv(out_path, header=False, index=False)

def pdwrite(in_path, out_path):
    df = pd.read_csv(in_path, sep='\t', header=None, dtype=object)
    df.to_excel(out_path, header=False, index=False)

if __name__ == "__main__":
    import sys
    if sys.argv[1] == '-r':
        pdread('App_Data/DataIn/am0411.xlsx', 'App_Data/DataOut/px.txt')
    if sys.argv[1] == '-w':
        pdwrite('App_Data/DataIn/am0411.txt', 'App_Data/DataOut/px.xlsx')