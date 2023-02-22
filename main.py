from convert import convert_xls_to_xlsx
from data_to_excel import writing_data_to_excel
from session import session2, header
from create_an_excel import data_frame
from const import CONST_URL, CONST_FEDEX_URL

print("started parsing...")
for i in range(5, 995):
    if i < 10:
        url = f"{CONST_URL}00{i}.xls"
        response = session2.get(url, headers=header, allow_redirects=True)

        if response.headers.get("content-type") == "application/vnd.ms-excel":
            open(f"xlsx/00{i}.xls", "wb").write(response.content)
            data_rows_for_sheet_2 = [[f"00{i}00-00{i}99", f"00{i}00", f"00{i}99"]]
            data_rows_for_sheet_1 = [[f"00{i}00-00{i}99", f"{CONST_FEDEX_URL}00{i}00-00{i}99.csv"]]
            writing_data_to_excel(data_rows_for_sheet_1, data_rows_for_sheet_2)

    elif 10 <= i < 100:
        url = f"{CONST_URL}0{i}.xls"
        response = session2.get(url, headers=header, allow_redirects=True)

        if response.headers.get("content-type") == "application/vnd.ms-excel":
            open(f"xlsx/0{i}.xls", "wb").write(response.content)
            data_rows_for_sheet_2 = [[f"0{i}00-0{i}99", f"0{i}00", f"0{i}99"]]
            data_rows_for_sheet_1 = [[f"0{i}00-0{i}99", f"{CONST_FEDEX_URL}0{i}00-0{i}99.csv"]]
            writing_data_to_excel(data_rows_for_sheet_1, data_rows_for_sheet_2)

    else:
        url = f"{CONST_URL}{i}.xls"
        response = session2.get(url, headers=header, allow_redirects=True)

        if response.headers.get("content-type") == "application/vnd.ms-excel":
            open(f"xlsx/{i}.xls", "wb").write(response.content)
            data_rows_for_sheet_2 = [[f"{i}00-{i}99", f"{i}00", f"{i}99"]]
            data_rows_for_sheet_1 = [[f"{i}00-{i}99", f"{CONST_FEDEX_URL}{i}00-{i}99.csv"]]
            writing_data_to_excel(data_rows_for_sheet_1, data_rows_for_sheet_2)

data_frame.save("Carriers zone ranges.xlsx")

print("parsing completed")
convert_xls_to_xlsx()
