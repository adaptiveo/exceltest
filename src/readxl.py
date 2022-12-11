import openpyxl as opxl

FILENAME = "../excel/test.xlsx"
USD = "USD"
EUR = "EUR"
RUR = "RUR"

# select workbook
wb = opxl.load_workbook(FILENAME, data_only=True)

# select worksheet
ws = wb.active

max_row_num = ws.max_row
min_row_num = ws.min_row

sum_usd = [0, 0]
sum_eur = [0, 0]
sum_rur = [0, 0]
shift = 0

for i in range(min_row_num,max_row_num+1):
    curr_cur_cell = ws.cell(i,3).value
    if curr_cur_cell == None:
        shift = 1
        continue
    curr_num_cell = ws.cell(i, 5).value
    if curr_cur_cell == USD:
        sum_usd[shift] = sum_usd[shift] + curr_num_cell
    elif curr_cur_cell == EUR:
        sum_eur[shift] = sum_eur[shift] + curr_num_cell
    elif curr_cur_cell == RUR:
        sum_rur[shift] = sum_rur[shift] + curr_num_cell

print(sum_usd)