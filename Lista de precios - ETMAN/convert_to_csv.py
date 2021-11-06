import pandas as pd


def conversion(filepath,filename):
    read_file = pd.read_excel(filepath)
    read_file.to_csv(f'{filename}.csv', sep=";", index=None,float_format = '%.2f', header=False, decimal=".")
    print("ok")
