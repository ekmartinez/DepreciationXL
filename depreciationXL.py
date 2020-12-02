import os, datetime, sys
import pandas as pd
import xlwings as xw
from xlwings import constants
from cashflows import cashflow, depreciation_sl

def header():
    '''Prints Header'''
    os.system('cls')
    print(f"\n{'^'*55}'\n        DEPRECIATION SCHEDULE GENERATOR\n{'˅'*55}")

def mainFunction():
    '''Requests data for processing'''
    loop_indicator = 1
    while loop_indicator == 1:
        header()
        asset = input('\nEnter Asset Description >>> ')
        cost = input('Enter cost >>> ')
        life = input('Enter depreciable life (in years) >>> ')
        date = input('Enter date aquired >>> ')

        info = {'Asset' : asset,
                'Cost' : float(cost),
                'Life' : int(life),
                'Date' : date,
                }

        print('\nAsset Description\n' + ('-'*55))
        for keys, values in info.items():
            print(f"{keys} : {values}")
        print('-'*55)

        ask = input("""
Please review the information provided, Do you wish to continue?
(Press Enter to continue, no to quit, try to try again) >>> """)
        print('')

        if ask.lower() == '':
            dep_rep_sl(info)
            xl_format(info)
            ask_continue = input("""
Operation Completed!
Press Enter to continue, type exit to quit program >>> """)
            if ask_continue == '':
                continue
            elif ask_continue == 'exit':
                loop_indicator = 0
            else:
                input('Invalid answer, returning to main program.')

        elif ask.lower() == 'no':
            input('\nPress enter to quit...')
            break
        elif ask.lower() == 'try':
            continue
        else:
            input('Invalid answer, press Enter to try again.....')

def dep_rep_sl(data):
    '''Generates depreciation schedule and exports it to Excel'''
    costs = cashflow(const_value=0, periods=data['Life']*12, start=data['Date'], freq='M')
    life = cashflow(const_value=0, periods=data['Life']*12, start=data['Date'], freq='M')
    costs[0] = data['Cost']
    life[0] = data['Life']*12
    df = depreciation_sl(costs=costs, life=life, salvalue=None)
    print(df) 

    ask_export = input("""
Do you wish to export the results to an Excel Workbook?
(Enter yes or no >>> """)
    xlFilename = f"{data['Asset']}_depreciation_schedule.xlsx"
    while(True):
        if ask_export.lower() == 'yes':
            df.to_excel(xlFilename)
            break
        elif ask_export.lower() == 'no':
            input('Press Enter to exit...')
            sys.exit()
        else:
            print('Invalid answer, try again.')

def xl_format(data):
        '''Formats exported Excel Document'''
        xlFilename = f"{data['Asset']}_depreciation_schedule.xlsx"
        wb = xw.Book(xlFilename)
        sht = wb.sheets('Sheet1')
        sht.range('A:F').api.Font.Bold = False
        sht.range('1:10').api.Insert()
        sht.range('A2').value = 'Depreciation Schedule Generator'
        sht.range('A2').api.Font.Bold = True
        sht.range('A2:F2').api.MergeCells = True
        sht.range('A3:B3').value = ['Created by:', os.getlogin()]
        sht.range('A4:B4').value = ['Generated on:', datetime.date.today()]
        sht.range('A6:B6').value = ['Asset', data['Asset']]
        sht.range('A7:B7').value = ['Cost', data['Cost']]
        sht.range('A8:B8').value = ['Life', data['Life']]
        sht.range('A9:B9').value = ['Purchased', data['Date']]
        sht.range('A11').value = 'Date'
        sht.range('A11:F11').api.Font.Bold = True
        sht.range('A:F').columns.autofit()
        sht.range('B3:B9').api.HorizontalAlignment = constants.HAlign.xlHAlignRight

mainFunction()



