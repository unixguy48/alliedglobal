from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.FileSystem import FileSystem

import re

'''
RPA Challenge
Developed by: Iveen Duarte
'''

browser_lib = Selenium()
spreadsheet_lib = Files()
fs_lib = FileSystem()
department_selector = None

def open_url(url):
    global browser_lib
    try:
        browser_lib.open_available_browser(url)
        browser_lib.maximize_browser_window()
        browser_lib.logger.info(f'Opened {url} successfully.')
    except Exception as e:
        browser_lib.logger.fatal(f'Could not open {url}.')
        browser_lib.logger.fatal(e)

def dive_in(element):
    global browser_lib
    browser_lib.wait_until_element_is_visible(element,60)
    is_button_visible = browser_lib.is_element_visible(element)
    if is_button_visible:
        browser_lib.click_element(element)


def navigate_departments(dpt_name):
    global browser_lib
    global department_selector
    data = []
    item = {}

    for row in range(1, 10):
        if row == 9:
            max_col = 3
        else:
            max_col = 4
        for col in range(1, max_col):
            dept_selector = f'css:#agency-tiles-widget .row:nth-child({row}) > .col-sm-4:nth-child({col}) .h4'
            amt_selector = f'css:#agency-tiles-widget .row:nth-child({row}) > .col-sm-4:nth-child({col}) .h1'
            browser_lib.wait_until_element_is_visible(dept_selector, 60)
            department_name = browser_lib.get_text(dept_selector)
            browser_lib.wait_until_element_is_visible(amt_selector)
            department_amount = browser_lib.get_text(amt_selector)
            item["Department"] = department_name
            item["Amount"] = department_amount
            data.append(item.copy())
            item.clear()
            if dpt_name == department_name:
                department_selector = dept_selector

    return data

def create_spreadsheet(filename):
    global spreadsheet_lib
    global fs_lib
    if fs_lib.does_file_exist(filename):
        fs_lib.remove_file(filename)
    spreadsheet_lib.create_workbook(filename)

def write_spreadsheet(data, filename, sheetname):
    create_spreadsheet(filename)
    spreadsheet_lib.rename_worksheet('Sheet', sheetname)
    spreadsheet_lib.append_rows_to_worksheet(data, header=True)
    spreadsheet_lib.save_workbook()
    spreadsheet_lib.close_workbook()
    spreadsheet_lib.logger.info(f'Saved {filename}.')


def get_selected_department():
    global fs_lib
    txt = fs_lib.read_file('input/department.txt')
    return txt.strip()

def get_headers():
    global browser_lib
    headers = ['UII']
    for i in range(2,8):
        selector = f'css:tr:nth-child(2) > .sorting:nth-child({i})'
        browser_lib.wait_until_element_is_visible(selector, 60)
        header_text = browser_lib.get_text(selector)
        headers.append(header_text)
    return headers


def download_file(uii):
    global browser_lib
    global fs_lib
    download_path = 'C:\\Users\iveen\\Downloads\\'
    other_site = Selenium()
    url = browser_lib.get_location()
    new_url = f'{url}/{uii}'
    try:
        if fs_lib.does_file_exist(f'{download_path}{uii}.pdf') == False:
            other_site.open_available_browser(new_url)
            other_site.wait_until_element_is_visible('css:#business-case-pdf > a', 10)
            if other_site.is_element_visible('css:#business-case-pdf > a'):
                other_site.click_element('css:#business-case-pdf > a')
                uii = uii.replace('#','')
                fs_lib.wait_until_created(f'{download_path}{uii}.pdf')
                fs_lib.logger.info(f'File {uii}.pdf downloaded.')
        else:
            fs_lib.logger.info(f'File {uii}.pdf was already stored.')
    except Exception as e:
        browser_lib.logger.info(f"{uii} doesn't have a file attachment.")
    finally:
        other_site.close_browser()
        

def get_items(headers, items_count):
    global browser_lib
    output_list = []
    item = {}
    for row in range(1, items_count + 1):
        browser_lib.logger.info(f'Processing Item {row} of {items_count}')
        if row % 2 == 0:
            isEven = True
        else:
            isEven = False
        if isEven:
            uii_selector = f'css:.even:nth-child({row}) > .sorting_2' # UII
            browser_lib.wait_until_element_is_visible(uii_selector, 60)
            uii = browser_lib.get_text(uii_selector)
            selector = f'css:.even:nth-child({row}) > .left:nth-child(2)' #  Bureau
            browser_lib.wait_until_element_is_visible(selector)
            bureau = browser_lib.get_text(selector)
            selector = f'css:.even:nth-child({row}) > .left:nth-child(3)' #  Title
            browser_lib.wait_until_element_is_visible(selector)
            title = browser_lib.get_text(selector)
            selector = f'css:.even:nth-child({row}) > .right' #  Amount
            browser_lib.wait_until_element_is_visible(selector)
            amount = browser_lib.get_text(selector)
            selector = f'css:.even:nth-child({row}) > .left:nth-child(5)' #  Type
            browser_lib.wait_until_element_is_visible(selector)
            type = browser_lib.get_text(selector)
            selector = f'css:.even:nth-child({row}) > .center:nth-child(6)' #  Rating
            browser_lib.wait_until_element_is_visible(selector)
            rating = browser_lib.get_text(selector)
            selector = f'css:.even:nth-child({row}) > .center:nth-child(7)' #  Projects
            browser_lib.wait_until_element_is_visible(selector)
            projects = browser_lib.get_text(selector)
        else:
            uii_selector = f'css:.odd:nth-child({row}) > .sorting_2' # UII
            browser_lib.wait_until_element_is_visible(uii_selector, 60)
            uii = browser_lib.get_text(uii_selector)
            selector = f'css:.odd:nth-child({row}) > .left:nth-child(2)' #  Bureau
            browser_lib.wait_until_element_is_visible(selector)
            bureau = browser_lib.get_text(selector)
            selector = f'css:.odd:nth-child({row}) > .left:nth-child(3)' #  Title
            browser_lib.wait_until_element_is_visible(selector)
            title = browser_lib.get_text(selector)
            selector = f'css:.odd:nth-child({row}) > .right' #  Amount
            browser_lib.wait_until_element_is_visible(selector)
            amount = browser_lib.get_text(selector)
            selector = f'css:.odd:nth-child({row}) > .left:nth-child(5)' #  Type
            browser_lib.wait_until_element_is_visible(selector)
            type = browser_lib.get_text(selector)
            selector = f'css:.odd:nth-child({row}) > .center:nth-child(6)' #  Rating
            browser_lib.wait_until_element_is_visible(selector)
            rating = browser_lib.get_text(selector)
            selector = f'css:.odd:nth-child({row}) > .center:nth-child(7)' #  Projects
            browser_lib.wait_until_element_is_visible(selector)
            projects = browser_lib.get_text(selector)
        
        item[headers[0]] = uii
        item[headers[1]] = bureau
        item[headers[2]] = title
        item[headers[3]] = amount
        item[headers[4]] = type
        item[headers[5]] = rating
        item[headers[6]] = projects

        # download_file(uii)

        output_list.append(item.copy())
        item.clear()

    return output_list


def get_entries_count(text):
    records = '0'
    regex = r'Showing\s\d{1,4}\sto\s\d{1,4}\sof\s(\d{1,5})\sentries'
    if re.match(regex, text):
        m = re.match(regex, text)
        records = m.group(1)
    
    entries = int(records)
    return entries


def get_department_table():
    global browser_lib
    global department_selector
    
    all_items = []

    browser_lib.click_element(department_selector)
    browser_lib.wait_until_element_is_visible("xpath://div[@id='investments-table-object_wrapper']", 60)
    table_header = get_headers()
    browser_lib.wait_until_element_is_visible("css:#investments-table-object_info")
    entries_text = browser_lib.get_text("css:#investments-table-object_info")
    entries = get_entries_count(entries_text)
    browser_lib.wait_until_element_is_visible('css:.c-select:nth-child(1)')
    browser_lib.select_from_list_by_label('css:.c-select:nth-child(1)', 'All')
    all_items = get_items(table_header, entries)
    return all_items

def minimal_task():
    global department_selector
    department_name = get_selected_department()
    open_url('https://itdashboard.gov')
    print("Done.")
    dive_in("xpath://a[contains(.,'DIVE IN\u00a0')]")
    data = navigate_departments(department_name)
    write_spreadsheet(data, filename='output/Agencies_Expenditures.xlsx', sheetname='Agencies')
    items = get_department_table()
    write_spreadsheet(items, filename=f'output/{department_name}_Individual Investments.xlsx', sheetname='List')
    browser_lib.close_browser()


if __name__ == "__main__":
    minimal_task()
