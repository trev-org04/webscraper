from selenium import webdriver
from selenium.webdriver.common.by import By
from xlwt import Workbook, easyxf
from time import sleep

def createBankList():
    bank_list = []
    browser = webdriver.Chrome()
    browser.get('https://bank-code.net/routing-numbers')
    sleep(3.0)
    bank_columns = browser.find_elements(By.CLASS_NAME, 'col-xs-4')
    for col in bank_columns:
        banks = col.find_elements(By.TAG_NAME, 'a')
        for bank in banks:
            bank_list.append(bank.text)
    wb = Workbook()
    index = 0
    while index < len(bank_list) - 1:
        print(bank_list[index])
        if (len(bank_list[index]) > 31):
            bank_sheet_name = bank_list[index][:31]
            curr_sheet = wb.add_sheet(bank_sheet_name)
        else:
            curr_sheet = wb.add_sheet(bank_list[index])
        if (bank_list[index].find('-') != -1):
            bank_list[index] = bank_list[index].replace('-', '_')
        else:
            bank_list[index] = bank_list[index]      
        bank_list[index] = bank_list[index].replace(' ', '-').lower()
        if (bank_list[index].find(',') != -1):
            bank_web_name = bank_list[index].replace(',', '%2c')
        else:
            bank_web_name = bank_list[index]
        style = easyxf('font: bold 1')
        browser.get('https://bank-code.net/routing-numbers/bank/{}'.format(bank_web_name))
        sleep(3.0)
        has_next_page = len(browser.find_elements(By.CLASS_NAME, 'pagination')) > 0
        curr_sheet.write(0, 0, 'ROUTING NUMBER', style)
        curr_sheet.write(0, 1, 'ADDRESS', style)
        curr_sheet.write(0, 2, 'CITY', style)
        curr_sheet.write(0, 3, 'STATE', style)
        cells = browser.find_element(By.CLASS_NAME, 'table').find_elements(By.TAG_NAME, 'td')
        row_index = 1
        data_index = 0
        for cell in cells:
            if cell.text != '':
                if data_index == 1:
                    curr_sheet.write(row_index, 0, cell.text)
                elif data_index == 2:
                    curr_sheet.write(row_index, 1, cell.text)
                elif data_index == 3:
                    curr_sheet.write(row_index, 2, cell.text)
                elif data_index == 4:
                    curr_sheet.write(row_index, 3, cell.text)
                    data_index = -1
                    row_index += 1
                data_index += 1
        if has_next_page:
            browser.get('https://bank-code.net/routing-numbers/bank/{}/100'.format(bank_web_name))
            sleep(3.0)
            more_cells = browser.find_element(By.CLASS_NAME, 'table').find_elements(By.TAG_NAME, 'td')
            row_index = 101
            data_index = 0
            for cell in more_cells:
                if cell.text != '':
                    if data_index == 1:
                        curr_sheet.write(row_index, 0, cell.text)
                    elif data_index == 2:
                        curr_sheet.write(row_index, 1, cell.text)
                    elif data_index == 3:
                        curr_sheet.write(row_index, 2, cell.text)
                    elif data_index == 4:
                        curr_sheet.write(row_index, 3, cell.text)
                        data_index = -1
                        row_index += 1
                    data_index += 1
        index += 1
    wb.save('Bank List.xls')

createBankList()