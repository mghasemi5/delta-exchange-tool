from win32com.client import Dispatch
import os
import time

while True:

    try:



        xl = Dispatch('Excel.Application')
        wb_eth = xl.Workbooks.Open(r'C:\Users\mogh77\Desktop\delta_api\eth.xlsx')
        path_to_pdf_eth = r'C:\Users\mogh77\Desktop\delta_api\pdf\eth.pdf'
        wb_btc = xl.Workbooks.Open(r'C:\Users\mogh77\Desktop\delta_api\btc.xlsx')
        path_to_pdf_btc = r'C:\Users\mogh77\Desktop\delta_api\pdf\btc.pdf'
        wb_eth.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf_eth)
        wb_btc.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf_btc)
        print('done')
        time.sleep(30)


    except:
            print('rid')
            continue
