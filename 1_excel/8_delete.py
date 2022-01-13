from openpyxl import load_workbook

wb = load_workbook("sample.xlsx")
ws = wb.active

ws.delete_rows(8,3) #8번째 줄부터 총 3줄 삭제 (8,9,10행)
wb.save("sample_delete_row.xlsx")

ws.delete_cols(2,2) #2번째 열로부터 총 2개의 열 삭제 (B,C)
wb.save("sample_delete_cols.xlsx")
