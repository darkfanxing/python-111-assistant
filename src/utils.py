def write_head_data(sheet, row_index, data):
    for index, value in enumerate(data):
        sheet.write(row_index, index+1, value)