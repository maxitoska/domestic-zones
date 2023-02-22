import requests
from convert import convert_xls_to_xlsx
from create_an_excel import data_frame, sheet_2, sheet_1

const_url = "https://www.ups.com/media/us/currentrates/zone-csv/"
const_fedex_url = "https://www.fedex.com/ratetools/documents2/"
print("started parsing...")
for i in range(5, 995):
    if i < 10:
        url = f"{const_url}00{i}.xls"
        response = requests.get(url, allow_redirects=True)

        if response.headers.get("content-type") == "application/vnd.ms-excel":
            open(f"xls/00{i}.xls", "wb").write(response.content)
            data_rows_for_sheet_2 = [[f"00{i}00-00{i}99", f"00{i}00", f"00{i}99"]]
            data_rows_for_sheet_1 = [[f"00{i}00-00{i}99", f"{const_fedex_url}00{i}00-00{i}99.csv"]]
            # for sheet in data_frame.worksheets:
            for row in data_rows_for_sheet_2:
                sheet_2.append(row)
            for row in data_rows_for_sheet_1:
                sheet_1.append(row)
    elif 10 <= i < 100:
        url = f"{const_url}0{i}.xls"
        response = requests.get(url, allow_redirects=True)

        if response.headers.get("content-type") == "application/vnd.ms-excel":
            open(f"xls/0{i}.xls", "wb").write(response.content)
            data_rows_for_sheet_2 = [[f"0{i}00-0{i}99", f"0{i}00", f"0{i}99"]]
            data_rows_for_sheet_1 = [[f"0{i}00-0{i}99", f"{const_fedex_url}0{i}00-0{i}99.csv"]]
            for row in data_rows_for_sheet_2:
                sheet_2.append(row)
            for row in data_rows_for_sheet_1:
                sheet_1.append(row)
    else:
        url = f"{const_url}{i}.xls"
        response = requests.get(url, allow_redirects=True)

        if response.headers.get("content-type") == "application/vnd.ms-excel":
            open(f"xls/{i}.xls", "wb").write(response.content)
            data_rows_for_sheet_2 = [[f"{i}00-{i}99", f"{i}00", f"{i}99"]]
            data_rows_for_sheet_1 = [[f"{i}00-{i}99", f"{const_fedex_url}{i}00-{i}99.csv"]]

            for row in data_rows_for_sheet_2:
                sheet_2.append(row)
            for row in data_rows_for_sheet_1:
                sheet_1.append(row)

data_frame.save("Carriers zone ranges.xlsx")

print("parsing completed")
convert_xls_to_xlsx()
