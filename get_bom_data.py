import time
import pygetwindow as gw
import pyautogui as pygui
from coordinates import *

# fetch all bom data from misys through the menus
def get_bom_data():
    print("Getting data")
    pygui.click(file_menu)
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
    while len(gw.getWindowsWithTitle('Export Bills of Material Information')) == 1:
        time.sleep(1)

if __name__ == '__main__':
    get_bom_data()