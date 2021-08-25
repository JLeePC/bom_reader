import os
import openpyxl
import pyautogui as pygui
import pandas as pd
import datetime as dt
import pygetwindow as gw
import time

# coordinates
item_no = 250,94
rev_no = 458,94
header = 41,165
material = 207,163
line_top = 1885,235
line_up = 1884,255
line_down = 1885,276
line_bottom = 1884,300
new_line = 1886,322
delete_line = 1885,342
save = 76,58
file = 15,32
export = 54,185
export_type = 1129,357
path = 1104,383
deselect_bom = 771,437
deselect_bom_routing = 771,453
deselect_bom_location = 771,485
deselect_alt_comp = 771,500
export_ok = 781,701
select_csv = 880,396
export_complete_ok = 793,601

def get_bom_data():
    print("Getting data")
    pygui.click(file)
    pygui.click(export)
    pygui.click(export_type)
    pygui.typewrite(["down"])
    pygui.hotkey('enter')
    pygui.click(path)
    pygui.typewrite('D:\MIsys Data\Bill of Materials\Bill of Material Details.CSV')
    pygui.click(deselect_bom)
    pygui.click(deselect_bom_routing)
    pygui.click(deselect_bom_location)
    pygui.click(deselect_alt_comp)
    pygui.click(export_ok)
    while len(gw.getWindowsWithTitle('Export Bills of Material Information')) == 0:
        time.sleep(1)
    pygui.click(export_complete_ok)

def main():
    # see if the 'Bill of Materials.CSV' file exists and is recent. if its not export the data again
    today = dt.datetime.now().date()
    try:
        filetime = dt.datetime.fromtimestamp(os.path.getctime("D:\MIsys Data\Bill of Materials\Bill of Material Details.CSV"))

        if filetime.date() != today:
            get_bom_data()
        else:
            data = pd.read_csv('D:\MIsys Data\Bill of Materials\Bill of Material Details.CSV')
            active_bom = []
            for i in range(len(data)):
                builder = {}
                bom = data['bomItem'][i]
                rev = data['bomRev'][i]
                if "QS-101-2-3" in bom and '2' in rev:
                    builder['line'] = data['lineNbr'][i]
                    builder['part'] = data['partId'][i]
                    builder['qty'] = data['qty'][i]
                    builder['bom'] = bom
                    builder['rev'] = rev
                    active_bom.append(builder)
            ordered_bom = []
            for line in range(len(active_bom)):
                for i in active_bom:
                    if int(i['line']) == line - 1:
                        ordered_bom.append(i)
            for i in ordered_bom:
                print(i)

    except FileNotFoundError:
        print("'Bill of Material Details.CSV' File not found")
        get_bom_data()

if __name__ == "__main__":
    main()

print("Done.")