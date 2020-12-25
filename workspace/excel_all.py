from import_list import *

def len_byte(value):
    length = len(value)
    utf8_length = len(value.encode('utf-8'))
    length = (utf8_length - length) / 2 + length
    # print("value:", value, flush=True)
    # print("len_byte:", length,flush=True)
    # print("")
    return int(length)

def write_to_excel_all():

    results = gl.get_value('results')

    workbook = Workbook()
    sheet = workbook.active
    for num, result in enumerate(results):
        _, pic_name = os.path.split(result['save_path'])
        if num != 0:
            sheet = workbook.create_sheet()
            sheet.title = str(pic_name)
        else:
            sheet.title = str(pic_name)
        data = result['data']
        col_width = [0, 0]
        col_width[0] = len_byte('text')
        col_width[1] = len_byte('text_box_position')
        for pos, infomation in enumerate(data):
            if col_width[0] < len_byte(str(infomation['text'])):
                col_width[0] = len_byte(str(infomation['text']))
            if col_width[1] < len_byte(str(infomation['text_box_position'])):
                col_width[1] = len_byte(str(infomation['text_box_position']))
        if(gl.get_value('auto_width')==True):
            sheet.column_dimensions["A"].width = (col_width[0] + 1)
            sheet.column_dimensions["B"].width = (col_width[1] + 1)
        else:
            pass

        sheet.cell(row=1, column=1).value="text"
        sheet.cell(row=1, column=2).value="text_box_position"
        for pos, infomation in enumerate(data):
            sheet.cell(row=(pos+2), column=1).value=str(infomation['text'])
            sheet.cell(row=(pos+2), column=2).value=str(infomation['text_box_position'])

        if(gl.get_value('frozen')==True):
            sheet.freeze_panes = 'A2'
        else:
            sheet.freeze_panes = 'A1'

    workbook.save("C:/Users/m1861/Desktop/SCD/Output_Excel/" +"test" + ".xlsx") # 保存

if __name__ == "__main__":
    gl._init()
    results_demo()
    write_to_excel_all()