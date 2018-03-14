from openpyxl import load_workbook, Workbook
from openpyxl.styles import Border, Side, Font, Alignment, PatternFill, Color

def srs2xl(reqIdDic):
    wb=Workbook()
    ws1 = wb.active
    ws1.title = "Input_Sheet"
    ws1["A1"] = "Word"
    ws1["B1"] = "ReqID"
    for i in range(2,len(reqIdDic)) :
        for j in range(1,3) :
            c = ws1.cell(i,j)
            if reqIdDic[i-2][1] != None :
                c.value = reqIdDic[i-2][j-1]
    wb.save('출력결과.xlsx')

def srs_out(srs):
    pass