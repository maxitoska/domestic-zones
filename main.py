import requests
from convert import convert_xls_to_xlsx
from create_an_excel import data_frame

const_url = "https://www.ups.com/media/us/currentrates/zone-csv/"
for i in range(5, 995):
    if i < 10:
        url = f"{const_url}00{i}.xls"
        response = requests.get(url, allow_redirects=True)

        if response.headers.get('content-type') == "application/vnd.ms-excel":
            open(f'00{i}.xls', 'wb').write(response.content)
            data_rows = [[f"00{i}00-00{i}99", f"00{i}00", f"00{i}99"]]

            for sheet in data_frame.worksheets:
                for row in data_rows:
                    sheet.append(row)

    elif 10 <= i < 100:
        url = f"{const_url}0{i}.xls"
        response = requests.get(url, allow_redirects=True)

        if response.headers.get('content-type') == "application/vnd.ms-excel":
            open(f'0{i}.xls', 'wb').write(response.content)
            data_rows = [[f"0{i}00-0{i}99", f"0{i}00", f"0{i}99"]]

            for sheet in data_frame.worksheets:
                for row in data_rows:
                    sheet.append(row)

    else:
        url = f"{const_url}{i}.xls"
        response = requests.get(url, allow_redirects=True)

        if response.headers.get('content-type') == "application/vnd.ms-excel":
            open(f'{i}.xls', 'wb').write(response.content)
            data_rows = [[f"{i}00-{i}99", f"{i}00", f"{i}99"]]

            for sheet in data_frame.worksheets:
                for row in data_rows:
                    sheet.append(row)

        data_frame.save("Carriers zone ranges.xlsx")

convert_xls_to_xlsx()
