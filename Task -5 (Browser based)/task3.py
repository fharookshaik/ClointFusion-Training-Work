import ClointFusion as cf
import subprocess as sp 

cf.OFF_semi_automatic_mode()



def write_into_NameBox(inp):
    cf.mouse_click(x=73,y=224,single_double_triple='double')
    cf.key_write_enter(strMsg=inp,delay=0.5,key='e')



cf.window_show_desktop()
cf.message_counter_down_timer(strMsg='Starting Process',start_value=3)
cf.mouse_click(*cf.mouse_search_snip_return_coordinates_x_y('images/excel.png'),single_double_triple='double')



cf.message_counter_down_timer(strMsg='Looking for Table',start_value=3)

write_into_NameBox(inp='G1')
cf.key_press(strKeys='ctrl + down')
cf.key_press(strKeys='ctrl + shift + down')
cf.key_press(strKeys='ctrl + shift + left')
cf.key_press(strKeys='ctrl + c')
cf.key_press(strKeys='shift + f11')
cf.key_press(strKeys='ctrl + v')


cf.message_counter_down_timer(strMsg='Sorting the Region to central',start_value=3)

write_into_NameBox(inp='B1')
cf.key_press(strKeys='ctrl + shift + l')
cf.key_press(strKeys='alt + down')
cf.mouse_click(*cf.mouse_search_snip_return_coordinates_x_y('images/west.png'),single_double_triple='single')
cf.mouse_click(*cf.mouse_search_snip_return_coordinates_x_y('images/east.png'),single_double_triple='single')
cf.key_press(strKeys='enter')
cf.key_press(strKeys='ctrl + a')
cf.key_press(strKeys='ctrl + c')
cf.key_press(strKeys='shift + f11')
cf.key_press(strKeys='ctrl + v')


# cf.key_press('ctrl + a')
cf.message_counter_down_timer(strMsg='Adding Border to the Table',start_value=3)

cf.key_press('alt + h + b + a + enter')


cf.message_counter_down_timer(strMsg='Sorting the Total',start_value=3)

write_into_NameBox('G1')
cf.key_press(strKeys='alt + a + s + a + enter')
cf.key_hit_enter()



cf.message_counter_down_timer(strMsg='Changing the heading colors',start_value=3)

write_into_NameBox('A1')
cf.key_press('ctrl + shift + right')
cf.key_press('alt + h')
cf.key_press('alt + h + m')
cf.key_press('right')
cf.mouse_click(x=916,y=698,single_double_triple='double')
cf.key_write_enter(strMsg='000000',delay=0.5,key='e')


cf.key_press('alt + h + f + c + m')
cf.key_press('right')
cf.mouse_click(x=916,y=698,single_double_triple='double')
cf.key_write_enter(strMsg='FF9900',delay=0.5,key='e')



cf.message_counter_down_timer(strMsg='Conditional formatting the Total',start_value=3)

write_into_NameBox('G1')
cf.key_press('ctrl + shift + down')
cf.key_press('alt + h + l + s')
cf.key_press('alt + ctrl + enter')
write_into_NameBox('A1')

cf.key_press('ctrl + s')
cf.key_press('alt + f4')
cf.message_flash(msg=f'Task 3 Done {cf.show_emoji()}',delay=3)


cf.message_counter_down_timer("Calling & Executing task4.py",start_value=3)

sp.Popen('py task4.py',shell=True)