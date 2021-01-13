# -*- coding: utf-8 -*-

from bs4 import BeautifulSoup
import requests
from selenium import webdriver
import xlsxwriter
from time import sleep
from selenium.common.exceptions import NoSuchElementException

import re

# 크롬 드라이버 가져오기
driver = webdriver.Chrome('./chromedriver')

# 엑셀 파일 지정 & 쉬트지 열기
workbook = xlsxwriter.Workbook('Hmall_info_data.xlsx')
worksheet = workbook.add_worksheet()

f = open("url.txt", "r")
lines = f.read().split('\n')

#나중에 쓸 셀 형식들 (색상, 보더, 글 색상 등 )
data_format_first_line = workbook.add_format({'bg_color': '#E8F8FF', 'border': 7, 'bottom': 2, 'bold': True})
# data_format1 = workbook.add_format({'bg_color': '#E8F8FF', 'border': 7, 'bottom': 1})

#첫 줄 내용
row_one_line = ["상품 URL", "상품명"]
#첫 줄 형식지정 & 틀 고정
worksheet.set_row(0, cell_format=data_format_first_line)
worksheet.freeze_panes(1, 1)

row = 1
ul_row = 1

for url in lines:
    # 페이지로 가자p
    print(url)
    driver.get(url)
    print("go url page")
    # bs4
    req = requests.get(url) 
    html = req.text
    bs = BeautifulSoup(html, 'html.parser') 
    sleep(1)

    try:
        ele_names = driver.find_elements_by_class_name('pl_item_title')
    except NoSuchElementException:
        print("제목을 가져오는데 하는데 문제가 발생했습니다.")

    try:
        for ultag in bs.find_all('ul', {'class': 'pl_itemlist _4col'}):
            for litag in ultag.find_all('li'):
                ele_code = litag.attrs["id"]
                print (ele_code)
                worksheet.write(ul_row, 2, ele_code)
                ul_row = ul_row + 1
    except NoSuchElementException:
        ultag = "ele_model 직접 확인해봐!"
        print("ultag 문제가 발생했습니다.")


    for ele_name  in ele_names:
        worksheet.write(row, 1, ele_name.text)

        print(ele_name.text)
        row = row + 1



    # worksheet.write(row, 2, ele_price)
    # worksheet.write(row, 3, ele_maker)
    # worksheet.write(row, 4, ele_info)
    # worksheet.write(row, 5, ele_kc)
    # worksheet.write(row, 6, ele_return)
    # worksheet.write(row, 7, ele_model_two)

#첫 줄에 데이터 입력
for n, info in enumerate(row_one_line):
    worksheet.write(0, n, info)

workbook.close()
driver.close()
