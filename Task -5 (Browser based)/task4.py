import ClointFusion as cf
import os

cf.OFF_semi_automatic_mode() 


def getData(excelPath,sheetName,column_name):
    row, column = cf.excel_get_row_column_count(excel_path=excelPath,sheet_name=sheetName)
    data = []
    for i in range(0,row-1):
        temp = cf.excel_get_single_cell(excel_path=excelPath,sheet_name='Sheet1',columnName=column_name,cellNumber=i)
        if str(temp) != 'nan':
            data.append(temp)

    return data

def writeData(imgPath,data):
    cf.mouse_click(*cf.mouse_search_snip_return_coordinates_x_y(imgPath), single_double_triple='single')
    for i in data:
        cf.key_write_enter(strMsg=i,delay=0.5,key='e')

cwd = os.getcwd()

row,column = cf.excel_get_row_column_count(excel_path='bot_email_inputs.xlsx',sheet_name='Sheet1')
print(row, column)

email_to = getData(excelPath='bot_email_inputs.xlsx',sheetName='Sheet1',column_name='TO')
print(email_to)
cc = getData(excelPath='bot_email_inputs.xlsx',sheetName='Sheet1',column_name='CC')
print(cc)
subject = getData(excelPath='bot_email_inputs.xlsx',sheetName='Sheet1',column_name='Subject')
print(subject)
msg = getData(excelPath='bot_email_inputs.xlsx',sheetName='Sheet1',column_name='Message')
print(msg)
signature = getData(excelPath='bot_email_inputs.xlsx',sheetName='Sheet1',column_name='Signature')
print(signature)


cf.window_show_desktop()
cf.mouse_click(*cf.mouse_search_snip_return_coordinates_x_y('images/excel.png'), single_double_triple='double')
cf.key_press(strKeys='ctrl + a + c')


cf.window_show_desktop()
cf.key_press('windows')
cf.key_write_enter(strMsg='Office',delay=0.5,key='e')
cf.window_activate_and_maximize_windows('Office')
cf.mouse_click(*cf.mouse_search_snip_return_coordinates_x_y('images/outlook.png'), single_double_triple='single')
# cf.launch_any_exe_bat_application('Office')


cf.mouse_click(*cf.mouse_search_snip_return_coordinates_x_y('images/new_message.png'), single_double_triple='single')

writeData(imgPath='images/to.png',data=email_to)

cf.mouse_click(*cf.mouse_search_snip_return_coordinates_x_y('images/CC.png'), single_double_triple='single')
writeData(imgPath='images/cc_inp.png',data=cc)

writeData(imgPath='images/subject.png',data=subject)

writeData(imgPath='images/message.png',data=msg)

cf.key_press(strKeys='ctrl + v')

for i in signature:
    cf.key_write_enter(strMsg=i,delay=0.5,key='e')

cf.mouse_click(*cf.mouse_search_snip_return_coordinates_x_y('images/attach.png'), single_double_triple='single')


cf.mouse_click(*cf.mouse_search_snip_return_coordinates_x_y('images/upload.png'), single_double_triple='single')

cf.mouse_click(*cf.mouse_search_snip_return_coordinates_x_y('images/documents.png'), single_double_triple='single')
cf.mouse_click(*cf.mouse_search_snip_return_coordinates_x_y('images/attachment_path.png'), single_double_triple='single')
cf.key_write_enter(strMsg=cwd,delay=0.5,key='e')

cf.mouse_click(*cf.mouse_search_snip_return_coordinates_x_y('images/excel_file.png'), single_double_triple='single')
cf.key_hit_enter()

cf.mouse_click(*cf.mouse_search_snip_return_coordinates_x_y('images/send.png'), single_double_triple='single')


cf.message_flash(msg=f'Task 4 Done {cf.show_emoji()}',delay=3)


cf.message_flash(msg=f'All Done {cf.show_emoji()}',delay=3)