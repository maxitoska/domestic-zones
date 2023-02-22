import pandas as pd
import os
import glob


def convert_xls_to_xlsx():
    input_dir = f"{os.getcwd()}/xlsx"

    print("Started convert...")

    files = glob.glob(input_dir + "/*xls")
    for filename in files:
        out_name = f"{os.path.splitext(filename)[0]}.xlsx"
        if os.path.isfile(f"{input_dir}/{filename}x"):
            os.remove(f"{input_dir}/{filename}x")
        if filename.endswith(".xls") and out_name not in files:
            df = pd.read_excel(filename)
            df.to_excel(out_name)
            os.remove(filename)
    print("convert completed")
