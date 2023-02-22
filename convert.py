import win32com.client
import os
import glob


def convert_xls_to_xlsx():
    o = win32com.client.Dispatch("Excel.Application")
    o.Visible = False
    input_dir = r"D:\Test Project Dexima\domestic_zones"
    output_dir = r"D:\Test Project Dexima\domestic_zones\xlsx"
    files = glob.glob(input_dir + "/*.xls")

    for filename in files:
        file = os.path.basename(filename)
        output = output_dir + '/' + file.replace('.xls', '.xlsx')
        wb = o.Workbooks.Open(filename)
        wb.ActiveSheet.SaveAs(output, 51)
        wb.Close(True)
        os.remove(file)
