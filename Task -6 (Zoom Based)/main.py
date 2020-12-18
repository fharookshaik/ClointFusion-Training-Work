import ClointFusion as cf
import os
import shutil
from configparser import ConfigParser

cf.OFF_semi_automatic_mode()


def read_config(filename='config.ini',section='outlook'):
    parser = ConfigParser()
    parser.read(filename)
    
    details = {}
    if parser.has_section(section):
        items = parser.items(section)
        for item in items:
            details[item[0]] = item[1]

        print(details)

    else:
        raise Exception('{0} not found in {1} file'.format(section,filename))

    return details

def getData(excelPath='',sheet_name='',columnName=''):
    data = []
    row, column =cf.excel_get_row_column_count(excel_path=excelPath,sheet_name=sheet_name)
    for i in range(row-1):
        temp = cf.excel_get_single_cell(excel_path=excelPath,sheet_name=sheet_name,columnName=columnName,cellNumber=i)
        if str(temp) != 'nan':
            data.append(temp)
    return data

    
def set_hackathonVersion_date(msg,hackathon_version,date):
    cf.message_counter_down_timer('Setting up the hackathon version and date from the text file',start_value=3)
    keyword = 'Hackathon'
    dateKeyword = 'Time:'
    for i in range(len(msg)):
        line = msg[i]
        tempList = line.split()
        if keyword in tempList:
            index = tempList.index(keyword) + 1
            tempList[index] = str(hackathon_version[0])
            msg[i] = ' '.join(tempList)
            
        if dateKeyword in tempList:
            index = tempList.index(dateKeyword)
            for k in range(3):
                tempList.pop(index + 1)
            tempList[index+1] = date[0]
            msg[i] = ' '.join(tempList)
    return msg

def createDir(filePath):
    cf.message_counter_down_timer('Creating a new directory named "textfiles" in CWD',start_value=2)
    if os.path.isdir(filePath):
        shutil.rmtree(filePath)
        os.mkdir(filePath)
    else:
        os.mkdir(filePath)

def create_msg(msg,mailId,name):
    cf.message_counter_down_timer(f'Creating Invitation Message for {name}',start_value=2)
    tempList = msg[0].split()
    tempList[1] = f'{name},'
    msg[0] = ' '.join(tempList)
    # print(msg)

    filePath = f'{os.getcwd()}\\textfiles'

    with open(f'{name}.txt','w') as fp:
        fp.writelines(msg)
        
    fp.close()
    shutil.move(os.getcwd() + f'\\{name}.txt', filePath)

def setupMail():
    cf.message_counter_down_timer('Setting Up mail',start_value=3)
    details = read_config()
    email = details['emailid']
    password = details['password']
    # email = cf.gui_get_any_input_from_user(msgForUser='Email: ')
    # password = cf.gui_get_any_input_from_user(msgForUser='Password',password=True)

    cf.launch_website_h('https://outlook.live.com/owa/')
    cf.browser_wait_until_h(text='Sign in')
    cf.browser_mouse_click_h(User_Visible_Text_Element='Sign in')
    cf.browser_write_h(Value=email,User_Visible_Text_Element='Email, phone, or Skype')
    cf.browser_mouse_click_h(User_Visible_Text_Element='Next')
    cf.browser_write_h(Value=password,User_Visible_Text_Element='Password')
    cf.browser_mouse_click_h(User_Visible_Text_Element='Sign in')
    cf.browser_wait_until_h(text='New Message')
    
    return True

def sendMail(name,mailId):
    cf.browser_mouse_click_h(User_Visible_Text_Element='New Message')
    cf.browser_wait_until_h(text='To')
    cf.browser_write_h(Value=mailId,User_Visible_Text_Element='To')
    cf.browser_write_h(Value='Hackathon Invitation mail TEST',User_Visible_Text_Element='Add a Subject')
    # cf.browser_locate_element_h(element='//*[@id="app"]/div/div[2]/div[1]/div[1]/div[3]/div[2]/div/div[3]/div[1]/div/div/div/div[1]/div[2]/div[1]')
    cf.mouse_click(x=1196,y=684,single_double_triple='single')
    with open(f'textfiles/{name}.txt','r') as fp:
        for line in fp.readlines():
            cf.key_write_enter(strMsg=line,delay=0.1,key=' ')

    fp.close()

    cf.browser_mouse_click_h(User_Visible_Text_Element='Send')



if __name__ == "__main__":
    row, column = cf.excel_get_row_column_count(excel_path='Email.xlsx',sheet_name='Sheet1')
    headers = cf.excel_get_all_header_columns(excel_path='Email.xlsx',sheet_name='Sheet1')
    mailId = getData(excelPath='Email.xlsx',sheet_name='Sheet1',columnName=headers[0])
    names = getData(excelPath='Email.xlsx',sheet_name='Sheet1',columnName=headers[1])
    hackathon_version = getData(excelPath='Email.xlsx',sheet_name='Sheet1',columnName=headers[2])
    date = getData(excelPath='Email.xlsx',sheet_name='Sheet1',columnName=headers[3])
    # print(mailId)
    # print(names)
    # print(hackathon_version)
    # print(date)
    with open(file='Email.txt',mode='r') as fp:
        msg = fp.readlines()
    
    fp.close()

    filePath = f'{os.getcwd()}\\textfiles'
    createDir(filePath)

    modifiedMsg = set_hackathonVersion_date(msg,hackathon_version,date)
    # print(modifiedMsg)
    for i in range(row-1):
        create_msg(modifiedMsg,mailId[i],names[i])

    
    if setupMail():
        for i in range(row-1):
            sendMail(names[i],mailId[i])

    cf.message_flash(msg=f'All Done {cf.show_emoji()}',delay=5)
    cf.message_counter_down_timer("Closing Browser",start_value=10)
    cf.browser_quit_h()