import os
import openpyxl
import pandas as pd
from itertools import islice

def get_filenames_in(path, file_extension=None):
    filenames = os.listdir(path)
    if file_extension:
        filenames = [filename for filename in filenames if filename.endswith(file_extension)]
    return filenames

def xlsx_to_dataframe(path):
    wb = openpyxl.load_workbook(filename)
    ws = wb.worksheets[0]
    data = ws.values
    cols = next(data) # 첫 번째 행을 cols로 생각
    data = [d for d in data]
    return pd.DataFrame(data, columns=cols)

file_names = get_filenames_in(path="./", file_extension="xlsx")
filename = file_names[0]
df = xlsx_to_dataframe(path=filename)

print(df)
print(df.columns)
print(df.dtypes)
