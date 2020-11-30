'''
Browser Based Training Task -2

Today's task is to order these 2 headsets from flipkart just go till the payment option (Don't Pay for this) and close the browser.

Note:- But not using Browser functions. Use images , mouse click and keyboard functions.

1:- boAt Airdopes 131 Bluetooth Headset
2:- boAt Airdopes 402 Bluetooth Headset
'''

'''
Important Things to be Noted:

*Still in development. May or May not work in your machine.

1. This Particular code is designed as per my PC needs.
2. To run this code successfully, You Should Have Edge or Firefox browser installed on your 
   PC and it's shortcut should be placed on Desktop.
3. Several functions related to 'window' are used, so please note down that it may or may  
   not work other than Windows Operating System.
4. As of now the current code will run on MICROSOFT EDGE browser. To run in firefox browser 
   you need to do the following steps:
            i. Comment down the line 76
            ii. Uncomment the line 79
5. You should've pre-logged into Flipkart and you should have one atlest address and one 
   card saved in your account.
6. Please don't do anything whill the code is executing as it needs some inforamtion 
   attained by screenscraping,from mouse coordinates and keyboardthings.
7. Since the following code didn't use any of the browser automation functions, it needs 
    some images to identify the mouse coordinates.
8. Make sure that the images folder is present within the current working directory.
9. Some repeated things are written by new functions. No need to worry about that.
10. While executing the code you'll prompted to select the newely opened tab/ browser. This 
    helps giving a better information to code to handle things carefully.


Feel Free to Connect me for any queries.

'''


import ClointFusion as cf
cf.OFF_semi_automatic_mode()

def return_coordinates(img):
    coordinates = cf.mouse_search_snip_return_coordinates_x_y(img,conf=0.8)
    return coordinates

def click_mouse(coordinates,single_double_triple):
    cf.mouse_click(x=coordinates[0],y=coordinates[1],single_double_triple=single_double_triple)
    print(f"Clicked on {coordinates} {single_double_triple} times")

def search_for_product_and_add_to_cart(productName,img):
    cf.mouse_find_highlight_click(searchText='Search for products, brands and more')

    cf.message_counter_down_timer(strMsg=f'Searching for {productName}',start_value=3)

    # Clicks on 'Search for products, brands and more'
    # cf.mouse_click(x=545,y=163) # works for 1080 screen, varies for other display
    cf.key_write_enter(strMsg=productName,delay=1,key='e')

    cf.message_counter_down_timer(strMsg=f'Looking for {productName}',start_value=3)

    coordinates = return_coordinates(img)
    click_mouse(coordinates,'single')

    # Adding to Cart
    cf.message_counter_down_timer(strMsg='Adding to Cart',start_value=3)
    coordinates = return_coordinates(img='images/add_to_cart.png')
    click_mouse(coordinates,'single')

if __name__ == "__main__":
    cf.window_show_desktop()
    # For edge browser
    coordinates = return_coordinates(img='images/edge.png')

    # for firefox browser
    # coordinates = return_coordinates(img='images/firefox.png')

    click_mouse(coordinates,'double')


    # Verify that the actual selected browser fire up
    cf.window_activate_and_maximize_windows(cf.gui_get_dropdownlist_values_from_user(msgForUser='Please Select the Newly Opened Browser from the list',dropdown_list=cf.window_get_all_opened_titles_windows(),multi_select=False)[0])

    cf.key_write_enter(strMsg='https://flipkart.com',key='e')
    cf.message_counter_down_timer(strMsg='Waiting until flipkart is open',start_value=3)

    # Project -1
    productName = 'boAt Airdopes 131 Bluetooth Headset'
    img = 'images/boat_airpodes_131_bluetooth_headset.png'
    search_for_product_and_add_to_cart(productName,img)

    # Product -2
    productName = 'boAt Airdopes 402 Bluetooth Headset'
    img = 'images/boat_airpodes_402_bluetooth_headset.png'
    search_for_product_and_add_to_cart(productName,img)


    # Place Order
    cf.message_counter_down_timer(strMsg='Placing Order',start_value=3)

    coordinates = return_coordinates(img='images/place_order.png')
    click_mouse(coordinates,'single')


    # Select the Address
    cf.message_flash(msg='Choose your address',delay=1000)

    coordinates = return_coordinates(img='images/deliver_here.png')
    click_mouse(coordinates,'single')


    # CVV Section
    cf.search_highlight_tab_enter_open(searchText='Continue')

    cf.message_flash(msg='Select Your Card and select OK. No Need to enter CVV',delay=1000)

    cvv = cf.gui_get_any_input_from_user(msgForUser='Enter cvv',password=True,mandatory_field=True)

    cf.key_write_enter(cvv)

    # Final Section

    # cf.message_counter_down_timer(strMsg='Aborting Transaction',start_value=3)
    cf.message_flash(msg='Please Abort Your Transaction for no loss',delay=10)
    # cf.search_highlight_tab_enter_open(searchText='Click Here to Abort Transaction')

    cf.message_counter_down_timer(strMsg='Exiting Browser')

    coordinates = return_coordinates('images/close.png')
    click_mouse(coordinates,'single')

    print("Success",cf.show_emoji())