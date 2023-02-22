import win32com.client
import os
import glob


def convert_xls_to_xlsx():
    print("Started convert...")
    o = win32com.client.Dispatch("Excel.Application")
    o.Visible = False
    input_dir = f"{os.getcwd()}/xls"
    output_dir = f"{os.getcwd()}/xlsx"
    files = glob.glob(input_dir + "/*.xls")

    for filename in files:
        file = os.path.basename(filename)
        output = output_dir + "/" + file.replace(".xls", ".xlsx")
        wb = o.Workbooks.Open(filename)
        wb.ActiveSheet.SaveAs(output, 51)
        wb.Close(True)
        os.remove(input_dir + "/" + file)

    print("convert completed")
