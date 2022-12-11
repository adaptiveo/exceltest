from dataclasses import dataclass
from typing import Any

import openpyxl as opxl

FILENAME = "../excel/test.xlsx"
USD = "USD"
EUR = "EUR"
RUR = "RUR"
CUR_NAME_ROW = 3
VAL_NAME_ROW = 5
COME = 0
OUT = 1

@dataclass
class CCell:
    type:str
    val:Any

# select workbook
wb = opxl.load_workbook(FILENAME, data_only=True)

# select worksheet
ws = wb.active

max_row_num = ws.max_row
min_row_num = ws.min_row

sum_usd = [0, 0]
sum_eur = [0, 0]
sum_rur = [0, 0]
shift = COME

for i in range(min_row_num,max_row_num+1):
    curr_cur_cell = ws.cell(i,CUR_NAME_ROW).value
    if curr_cur_cell == None:
        shift = OUT
        continue
    curr_num_cell = CCell (ws.cell(i, VAL_NAME_ROW).data_type, ws.cell(i, VAL_NAME_ROW).value)
    if curr_num_cell.type == "n":
        if curr_cur_cell == USD:
            sum_usd[shift] = sum_usd[shift] + curr_num_cell.val
        elif curr_cur_cell == EUR:
            sum_eur[shift] = sum_eur[shift] + curr_num_cell.val
        elif curr_cur_cell == RUR:
            sum_rur[shift] = sum_rur[shift] + curr_num_cell.val

print(sum_usd)