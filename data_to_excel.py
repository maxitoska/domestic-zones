from create_an_excel import sheet_2, sheet_1


def writing_data_to_excel(data_for_sheet_1: list, data_for_sheet_2: list):
    for row in data_for_sheet_2:
        sheet_2.append(row)
    for row in data_for_sheet_1:
        sheet_1.append(row)
