import ClointFusion as cf
import shutil
import os
import time
import subprocess as sp

cf.OFF_semi_automatic_mode()

cf.message_counter_down_timer('Starting Process',start_value=3)

cf.message_counter_down_timer("Opening Browser",start_value=3)

cf.launch_website_h(URL='https://www.contextures.com/xlsampledata01.html')

cf.message_counter_down_timer("Etracting data",start_value=3)

cf.key_press(strKeys='ctrl + a + c')
cf.key_press('alt + f4')

cf.message_counter_down_timer("Opening excel",start_value=3)

cf.key_press(strKeys='windows + r')
cf.key_write_enter(strMsg='excel',delay=3)
cf.key_hit_enter()
time.sleep(3)

cf.key_press('ctrl + v')
time.sleep(3)
cf.key_press('ctrl + s')
cf.key_write_enter(strMsg="Excel")
cf.key_press('alt + f4')

excelPath = cf.gui_get_any_file_from_user(msgForUser="Please select the newly excel file(usally it'll  be in Documents folder)",Extension_Without_Dot='xlsx')

shutil.move(src=excelPath,dst=os.getcwd())
shutil.copyfile(src= os.path.join(os.getcwd() , 'Excel.xlsx'),dst= os.path.join(os.environ['USERPROFILE'], 'Desktop','Excel.xlsx'))


cf.message_counter_down_timer("Calling & Executing task3 ",start_value=3)

sp.Popen('py task3.py',shell=True)
