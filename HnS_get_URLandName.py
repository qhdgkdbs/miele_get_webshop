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
workbook = xlsxwriter.Workbook('HNS_info.xlsx')
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
    # req = requests.get(url)
    # html = req.text
    # bs = BeautifulSoup(html, 'html.parser')
    # sleep(1)

    #
    # try:
    #     ele_names = bs.findAll("span", {"class": "nameTit"})
    #     for ele_name in ele_names:
    #         print(ele_name.text)
    #         worksheet.write(row, 2, ele_name.text)
    #         row = row + 1
    #
    #     row = 1
    #     ele_urls_p = bs.findAll("p", {"class": "goodsName"})
    #     for ele_url_p in ele_urls_p:
    #         print(ele_url_p.find("a")["href"])
    #         worksheet.write(row, 3, ele_url_p.find("a")["href"])
    #         row = row + 1
    # except NoSuchElementException:
    #     ele_name = "err"
    #     worksheet.write(row, 3, ele_name.text)
    #     print(ele_name.text)


    try:
        ele_price = driver.find_element_by_xpath('//*[@id="container"]/form/div[2]/div[2]/ul/li[1]/dl/dd/span/em').text
    except NoSuchElementException:
        ele_model = "err"
        print("price 문제가 발생했습니다.")

    try:
        ele_pd_name = driver.find_element_by_xpath('//*[@id="itemDetail1"]/table/tbody/tr[6]/td').text
    except NoSuchElementException:
        ele_pd_name = "err"
        print("cat 문제가 발생했습니다.")

    try:
        ele_code = driver.find_element_by_xpath('//*[@id="container"]/form/div[1]/dl/dd').text
    except NoSuchElementException:
        ele_code = "err"
        print("maker 문제가 발생했습니다.")

    try:
        ele_maker = driver.find_element_by_xpath('// *[ @ id = "itemDetail1"] / table / tbody / tr[1] / td').text
    except NoSuchElementException:
        ele_maker = "err"
        print("maker 문제가 발생했습니다.")

    try:
        ele_from = driver.find_element_by_xpath('//*[@id="itemDetail1"]/table/tbody/tr[2]/td').text
    except NoSuchElementException:
        ele_from = "err"
        print("from 문제가 발생했습니다.")

    try:
        ele_change = driver.find_element_by_xpath('//*[@id="container"]/form/div[2]/div[2]/ul/li[2]/dl[4]/dd').text
    except NoSuchElementException:
        ele_change = "err"
        print("change 문제가 발생했습니다.")

    try:
        ele_full_info = driver.find_element_by_xpath('//*[@id="itemDetail1"]/table').text
    except NoSuchElementException:
        ele_full_info = "err"
        print("fullInfo 문제가 발생했습니다.")

    try:
        ele_kc = driver.find_element_by_xpath('//*[@id="itemDetail1"]/table/tbody/tr[7]/td').text
    except NoSuchElementException:
        ele_kc = "err"
        print("kc 문제가 발생했습니다.")

    try:
        ele_cat = driver.find_element_by_xpath('// *[ @ id = "container"] / dl').text
    except NoSuchElementException:
        ele_cat = "err"
        print("kc 문제가 발생했습니다.")






    worksheet.write(row, 3, ele_price)
    worksheet.write(row, 4, ele_pd_name)
    worksheet.write(row, 5, ele_maker)
    worksheet.write(row, 6, ele_from)
    worksheet.write(row, 7, ele_change)
    worksheet.write(row, 8, ele_full_info)
    worksheet.write(row, 9, ele_kc)
    worksheet.write(row, 10, ele_code)
    worksheet.write(row, 11, ele_cat)

    row = row + 1

#첫 줄에 데이터 입력
for n, info in enumerate(row_one_line):
    worksheet.write(0, n, info)

workbook.close()
driver.close()
