import json
import uuid
import time
import socket
import traceback
import xlrd
import threading
from selenium import webdriver
from utils.StringUtils import contain
from selenium.webdriver.common.by import By
from queue import Queue
from utils.HubstudioUtils2 import startEnv, openClient


def nextTab(driver):
    tabs = driver.window_handles
    index = tabs.index(driver.current_window_handle)
    driver.switch_to.window(tabs[index + 1])


def do_run(table, ncols, queue):
    envname = table.cell_value(i, 0)
    print(envname)
    driver = startEnv(envname)
    for j in range(1, ncols):
        if table.cell_value(i, 0) != '':
            print(table.cell_value(i, j))
            driver.execute_script("window.open('" + table.cell_value(i, j) + "');")
            if contain(table.cell_value(i, j), ""):
                try:
                    nextTab(driver)
                    time.sleep(5)
                    ele = driver.find_element(By.XPATH, "//span[text()='Follow']")
                    ele.click()
                    time.sleep(1)
                except Exception as err:
                    print(traceback.print_exc())
                    print(err)

    # closeEnv(envname)
    queue.get()


if __name__ == '__main__':
    openClient()
    xlsx = xlrd.open_workbook("open.xlsx")
    sheet = xlsx.sheet_by_index(0)
    nrows = sheet.nrows
    print(nrows)
    ncols = sheet.ncols
    print(ncols)
    q = Queue(maxsize=1)
    for i in range(1, nrows):
        q.put(i)
        threading.Thread(target=do_run, args=(sheet, ncols, q), name=str(i)).start()
        time.sleep(2)
