import xlsxwriter
from RPA.Browser.Selenium import Selenium
import time
from bs4 import BeautifulSoup
import pdfplumber
from os import listdir
from os.path import isfile, join
import pandas as pd

browser_lib = Selenium()
browser_lib.set_download_directory('output')
workbook = xlsxwriter.Workbook('output/Agencies.xlsx')
worksheet = workbook.add_worksheet('Agencies')

def open_the_website(url):
    browser_lib.open_available_browser(url)

def find_text(term):
    return browser_lib.get_text(term)

def press_button(term):
    browser_lib.press_keys(term, "ENTER")

def main():
    try:
        open_the_website('https://itdashboard.gov/')
        press_button("class:icon-guage")
        time.sleep(5)
        agencies = find_text("class:wrapper").splitlines()[0::4]
        spending = find_text("class:wrapper").splitlines()[2::4]
        i = 1
        for item in agencies:
            worksheet.write('A' + str(i), item)
            i = i + 1
        i = 1
        for item in spending:
            worksheet.write('B' + str(i), item)
            i = i + 1
        worksheet2 = workbook.add_worksheet('National_Science_Foundation')
        open_the_website('https://itdashboard.gov/drupal/summary/422')
        time.sleep(10)
        browser_lib.select_from_list_by_value("name:investments-table-object_length", "-1")
        time.sleep(10)
        code = browser_lib.get_element_attribute("class:dataTables_scrollBody", "outerHTML")
        soup = BeautifulSoup(code, 'html.parser')
        target = soup.find_all('tr')
        worksheet2.write('A1', 'UII')
        worksheet2.write('B1', 'Bureau')
        worksheet2.write('C1', 'Investment Title')
        worksheet2.write('D1', 'Total FY2021 Spending ($M)')
        worksheet2.write('E1', 'Type')
        worksheet2.write('F1', 'CIO Rating')
        worksheet2.write('G1', '# of Projects')
        worksheet2.write('H1', 'Href')
        i = -1
        for item in target:
            res = []
            for elem in item.find_all('td', style=False):
                res.append(elem.text)
                if 'href' in str(elem):
                    url = 'https://itdashboard.gov' + str(elem).split('"')[3]
                    res.append(url)
                    open_the_website(url)
                    time.sleep(3)
                    browser_lib.click_element("id:business-case-pdf")
                    time.sleep(10)
            try:
                if len(res) > 7:
                    res.append(res[1])
                    res.pop(1)
            except:
                pass  # without href
            for idx, month in enumerate(res):
                worksheet2.write(i, idx, res[idx])
            i = i + 1
        workbook.close()
    finally:
        browser_lib.close_all_browsers()



def parse_pdf():
    df = pd.ExcelFile('output/Agencies.xlsx').parse('National_Science_Foundation')
    uii=[]
    for elem in df['UII']:
        uii.append(elem)
    investment=[]
    for elem in df['Investment Title']:
        investment.append(elem)
    mypath = 'output'
    files = [f for f in listdir(mypath) if isfile(join(mypath, f))]
    for file in files:
        if file[-4:]=='.pdf':
            with pdfplumber.open(r'output/'+str(file)) as pdf:
                first_page = pdf.pages[0]
                list = first_page.extract_text().splitlines()
                for item in list:
                    if 'Name of this Investment' in item:
                        print (item.split(': ')[1] in investment)
                    elif 'Unique Investment Identifier (UII)' in item:
                        print (item.split(': ')[1] in uii)

if __name__ == "__main__":
    main()
    parse_pdf()
