import gspread
import pandas as pd
from openpyxl import load_workbook
import re
from selenium import webdriver
import os
from selenium.webdriver.common.by import By
import datetime
import time

Zero = True

while Zero:

    def Main():
        print('--------------------------------------------------')
        Menu_2 = input(
            '\n\nPress [1] to add asset info\nPress [2] to read all assets\nPress [3] to search by asset tag\nPress [4] to exit \n--> ')






        def Part_Entry():
            count = int (input ('How many entries are you making\n--> '))
            for i in range(count):

                run = 0
                print('--------------------------------------------------')
                print('Asset Information entry')
                print('--------------------------------------------------')
                sa = gspread.service_account(filename="printer-asset-finder-22ef6d3840cf.json")
                sh = sa.open("Printer Asset Tracker")
                # Sheet_title = input('Enter Day\n')
                try:
                    Asset_Tag = input('Asset Tag\n--> ')
                except:
                    pass
                try:
                    Cubical_Number = input('Cubical Number\n--> ')
                except:
                    pass
                try:
                    Model = input('Model\n--> ')
                except:
                    pass
                try:
                    User = input('User\n--> ')
                except:
                    pass
                try:
                    Floor = str(input('Floor\n--> '))
                except:
                    pass
                try:
                    now = datetime.datetime.now()
                    date = now.strftime("%Y-%m-%d %H:%M:%S")
                except:
                    pass

                print('--------------------------------------------------')

                sa = gspread.service_account(filename="printer-asset-finder-22ef6d3840cf.json")

                sh = sa.open("Printer Asset Tracker")
                ts = sh.worksheet('Printers')
                values_list = ts.row_values(1)
                values_list2 = ts.col_values(1)
                # print(ts.get_all_records())
                show = ts.get_all_records()

                for up in show:
                    # print(a)
                    key = list(up.keys())
                    value = list(up.values())
                    # print(key)
                    # print(value)
                    ## Validating to see if username and password are in the value variable

                    if Asset_Tag in value:
                        run = 1
                        break

                if run == 0:

                    sa = gspread.service_account(filename="printer-asset-finder-22ef6d3840cf.json")
                    sh = sa.open("Printer Asset Tracker")
                    # Sheet_title = input('Enter Day\n')

                    wks = sh.worksheet('Printers')
                    try:
                        df = pd.DataFrame(
                            {'Asset Tag': [Asset_Tag], 'Cubical': [Cubical_Number], 'Model': [Model], 'User': [User],
                             'Floor': [Floor], 'Date of last entry': [date]})
                    except:
                        pass
                    try:
                        df_values = df.values.tolist()
                    except:
                        pass
                    try:
                        sh.values_append('Printers', {'valueInputOption': 'RAW'}, {'values': df_values})
                    except:
                        pass

                    print('Your entry was entered successfully')


        def Sub_2():
            # read
            sa = gspread.service_account(filename="printer-asset-finder-22ef6d3840cf.json")
            sh = sa.open("Printer Asset Tracker")
            #print('test')
            Sheet_title = str("Printers")

            try:
                wks_2 = sh.worksheet(Sheet_title)
            except:
                pass
            # print(wks_2.get_all_records())
            try:
                for o in wks_2.get_all_records():
                    print(o)
            except:
                pass

        def Sub_3():

            Input_Number = input('Please enter The search criteria\n--> ')

            sa = gspread.service_account(filename="printer-asset-finder-22ef6d3840cf.json")
            sh = sa.open("Printer Asset Tracker")
            ts = sh.worksheet('Printers')
            values_list = ts.row_values(1)
            values_list2 = ts.col_values(1)
            # print(ts.get_all_records())
            show = ts.get_all_records()
            # print(show)

            for up in show:
                # print(a)
                key = list(up.keys())
                value = list(up.values())
                # print(value)
                # print(key)
                # print(value)
                ## Validating to see if username and password are in the value variable

                if Input_Number in value:
                    print(value)

        def Sub_4():
            Zero = False

        if Menu_2 == '1':
            Part_Entry()

        elif Menu_2 == '2':
            Sub_2()

        elif Menu_2 == '3':
            Sub_3()

        elif Menu_2 == '4':
            Sub_4()

    print('--------------------------------------------------')
    print('Printer Asset Tracker')
    print('--------------------------------------------------')

    ext = input('Press y to continue or x to exit\n--> ')

    if ext == 'y':
        Zero = True
        Main()
    elif ext == 'x':

        Zero = False
    else:

        Zero = False
