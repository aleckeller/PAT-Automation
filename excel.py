from openpyxl import load_workbook

def loadWorkbook(self):
    print "loading workbook..."
    wb = load_workbook('./patexcel.xlsx')
    print "workbook was loaded"
    print(wb.get_sheet_names())
