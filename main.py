import os
import openpyxl
import pandas as pd
from datetime import datetime
from copy import copy

def get_filenames_in(path, file_extension=None):
    filenames = os.listdir(path)
    if file_extension:
        filenames = [filename for filename in filenames if filename.endswith(file_extension)]
    return filenames

def convert2dataframe(wb):
    ws = wb.worksheets[0]
    data = ws.values
    cols = next(data) # 첫 번째 행을 cols로 생각
    data = [d for d in data]
    return pd.DataFrame(data, columns=cols)

def get_date_from_timestamp(timestamp):
    return datetime.strptime(str(timestamp), "%Y-%m-%d %H:%M:%S").date()

def replace_row_with(values, ws, rowidx, first_cells):
    print(f"Writing {rowidx} {values}...")
    for idx, value in enumerate(values):
        ws.cell(row=rowidx, column=idx+1).font = copy(first_cells[idx].font)
        ws.cell(row=rowidx, column=idx+1).border = copy(first_cells[idx].border)
        ws.cell(row=rowidx, column=idx+1).fill = copy(first_cells[idx].fill)
        ws.cell(row=rowidx, column=idx+1).number_format = copy(first_cells[idx].number_format)
        ws.cell(row=rowidx, column=idx+1).protection = copy(first_cells[idx].protection)
        ws.cell(row=rowidx, column=idx+1).alignment = copy(first_cells[idx].alignment)
        ws.cell(row=rowidx, column=idx+1).value = value

def save_dataframe_with_space(wb, df, filename, space):
    ws = wb.worksheets[0]
    
    rowidx = 2
    predate = get_date_from_timestamp(df.iloc[0, 0])
    first_cells = [ws.cell(row=2, column=1+i) for i in range(0, len(df.columns))] # 셀서식 유지를 위해 복사!
    for _, row in df.iterrows():
        curdate = get_date_from_timestamp(row.iloc[0])
        if curdate != predate:
            predate = curdate
            for i in range(0, space):
                replace_row_with([None]*len(row), ws, rowidx, first_cells)
                rowidx += 1
        replace_row_with(row.tolist(), ws, rowidx, first_cells)
        rowidx += 1

    wb.save(filename=f"seperated_{filename}")    

file_names = get_filenames_in(path="./", file_extension="xlsx")
filename = file_names[0]
wb = openpyxl.load_workbook(filename=filename)
df = convert2dataframe(wb)
save_dataframe_with_space(wb, df, filename, space=3)
wb.close()

