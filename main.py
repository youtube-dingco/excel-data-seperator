import os
import openpyxl
import pandas as pd
from datetime import datetime

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

def get_date_from_timestamp(timestamp):
    return datetime.strptime(str(timestamp), "%Y-%m-%d %H:%M:%S").date()

def save_dataframe_with_space(df, filename, space):
    wb = openpyxl.Workbook()
    ws = wb.worksheets[0]
    
    # colum 입력
    ws.append(df.columns.tolist())

    # data입력
    predate = get_date_from_timestamp(df.iloc[0, 0])
    for idx, row in df.iterrows():
        curdate = get_date_from_timestamp(row[0])
        if curdate != predate:
            predate = curdate
            for i in range(0, space):
                ws.append([])
        ws.append(row.tolist())

    wb.save(filename=f"seperated_{filename}")
    

file_names = get_filenames_in(path="./", file_extension="xlsx")
filename = file_names[0]
df = xlsx_to_dataframe(path=filename)

print(df)
print(df.columns)
print(df.dtypes)

save_dataframe_with_space(df, filename, space=3)