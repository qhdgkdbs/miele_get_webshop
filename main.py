from bs4 import BeautifulSoup
import requests
from selenium import webdriver
import xlsxwriter
from time import sleep
import re

# 크롬 드라이버 가져오기
driver = webdriver.Chrome('./chromedriver')

# 엑셀 파일 지정 & 쉬트지 열기
workbook = xlsxwriter.Workbook('web_shop_data.xlsx')
worksheet = workbook.add_worksheet()

WEB_SHOP_LOGIN_URL = "https://prod-live-miele-kr-azure.fse.intershop.de/INTERSHOP/web/WFS/Miele-Site/en_US/-/EUR/ViewApplication-Logout"
WEB_SHOP_ORDERS_PAGE = "https://prod-live-miele-kr-azure.fse.intershop.de/INTERSHOP/web/WFS/Miele-Site/en_US/KR/KRW/ViewOrderList_52-StartSearch?ChannelID=gBUKAB1b3SoAAAFcqzhnubUb"

ID = "krort"
PW = "intershop02"
ORG = "Miele"

#나중에 쓸 셀 형식들 (색상, 보더, 글 색상 등 )
data_format_first_line = workbook.add_format({'bg_color': '#E8F8FF', 'border': 7, 'bottom': 2, 'bold': True})
# data_format1 = workbook.add_format({'bg_color': '#E8F8FF', 'border': 7, 'bottom': 1})

#첫 줄 내용
row_one_line = ["구매자 이름", "주소", "Product ID", "Gross ID", "Net Total", "Shipping Cost", "Order Shipping Promotion", "Total", "Promotion Code"]

#첫 줄 형식지정 & 틀 고정
worksheet.set_row(0, cell_format=data_format_first_line)
worksheet.freeze_panes(1, 1)

# 로그인하자
driver.get(WEB_SHOP_LOGIN_URL)
driver.find_element_by_id("LoginForm_Login").send_keys(ID)
driver.find_element_by_id ("LoginForm_Password").send_keys(PW)
driver.find_element_by_id("LoginForm_RegistrationDomain").send_keys(ORG)
driver.find_element_by_class_name("loginbutton").click()
print("LOGIN")

sleep(1)

# 데이터 오더 페이지로 가자
driver.get(WEB_SHOP_ORDERS_PAGE)
print("go order page")

# 전체 취소 버튼을 눌러보자구
button = driver.find_element_by_xpath('//*[@id="order_status_values"]/tbody/tr[2]/td/a[2]').click()
print("click unselect all")

sleep(1)

# 라디오 버튼을 눌러~
button = driver.find_element_by_xpath('//*[@id="OrderStates_3"]').click()
print("click radio btn")

sleep(1)

#//*[@id="order_status_values"]/tbody/tr[3]/td/table[2]/tbody/tr/td[2]/input[3]
button = driver.find_element_by_xpath('//*[@id="order_status_values"]/tbody/tr[3]/td/table[2]/tbody/tr/td[2]/input[3]').click()
print("click find btn")

sleep(1)

# 50개 한번에 보기

button = driver.find_element_by_xpath('//*[@id="main_wrapper"]/tbody/tr/td/table/tbody/tr/td/form/table/tbody/tr[6]/td/table[2]/tbody/tr/td[2]/input[2]').click()
print("click 50 btn")

sleep(1)

# 정렬을 해보자구
button = driver.find_element_by_xpath('//*[@id="main_wrapper"]/tbody/tr/td/table/tbody/tr/td/form/table/tbody/tr[6]/td/table[1]/tbody/tr[1]/td[2]/a').click()
print("click find btn fir")

sleep(1)

button = driver.find_element_by_xpath('//*[@id="main_wrapper"]/tbody/tr/td/table/tbody/tr/td/form/table/tbody/tr[6]/td/table[1]/tbody/tr[1]/td[2]/a').click()
print("click find btn sec")

sleep(1)

# 이제 데이터를 가져와야해 ㅎ
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')

data_info_numbers = int(len(soup.find_all('a', class_='table_detail_link')))
print("total_1/3 data num : " + str(data_info_numbers/3))

url_datas = []
n=0

for i in range(0,data_info_numbers):
    # metadata = soup.find_all('div', class_='basicList_title__3P9Q7')[i]
    # title = metadata.a.get('title')
    # print("<제품명> : ", title)  # title
    if (i == (3*n)) :
        product_info_url = soup.find_all('a', class_='table_detail_link')[i]
        url_datas.insert(len(url_datas), product_info_url.get('href'))
        n = n + 1
    # print("<가격> : ", price)  # 가격

    # url = metadata.a.get('href')
    # # print("<url> : ", url)  # url

# print(url_datas)
print("가져온 URL 갯수 : " + str(len(url_datas)))

names = []
addresses = []
product_ids = []
gross_totals = []
net_totals = []
shipping_costs = []
shipping_promotions = []
totals = []

for n, one_line_url in enumerate(url_datas):
    driver.get(one_line_url)

    name = driver.find_element_by_xpath('//*[@id="main_wrapper"]/tbody/tr/td/table/tbody/tr/td/form/table/tbody/tr[7]/td[2]/a').text
    print("name")
    address = driver.find_element_by_xpath('//*[@id="main_wrapper"]/tbody/tr/td/table/tbody/tr/td/table[4]/tbody/tr/td').text
    print("address")
    product_id = driver.find_element_by_xpath('//*[@id="tableOrderDetails"]/tbody/tr[2]/td[2]/a').text
    print("product_id")
    gross_total = driver.find_element_by_xpath('//*[@id="tableOrderDetails"]/tbody/tr[2]/td[9]').text
    print("gross_total")
    net_total = driver.find_element_by_xpath('//*[@id="tableOrderDetails"]/tbody/tr[2]/td[6]').text
    print("net_total")
    shipping_cost = driver.find_element_by_xpath('//*[@id="main_wrapper"]/tbody/tr/td/table/tbody/tr/td/form/table/tbody/tr[23]/td/table/tbody/tr[4]/td[2]').text
    print("shipping_cost")
    shipping_promotion = driver.find_element_by_xpath('//*[@id="main_wrapper"]/tbody/tr/td/table/tbody/tr/td/form/table/tbody/tr[23]/td/table/tbody/tr[5]/td[2]').text
    print("shipping_promotion")
    # total = driver.find_element_by_xpath('//*[@id="main_wrapper"]/tbody/tr/td/table/tbody/tr/td/form/table/tbody/tr[23]/td/table/tbody/tr[13]/td[2]').text
    # print("total")




    names.insert(len(names), name)
    addresses.insert(len(addresses), address)
    product_ids.insert(len(product_ids), product_id)
    gross_totals.insert(len(gross_totals), gross_total)
    net_totals.insert(len(net_totals), net_total)
    shipping_costs.insert(len(shipping_costs), shipping_cost)
    shipping_promotions.insert(len(shipping_promotions), shipping_promotion)
    # totals.insert(len(totals), total)

start_row = 1
start_col = 0

for i in range(len(names)):
    worksheet.write(start_row, start_col, names[i])     #0
    start_col = start_col + 1
    worksheet.write(start_row, start_col, addresses[i]) #1
    start_col = start_col + 1
    worksheet.write(start_row, start_col, product_ids[i]) #2
    start_col = start_col + 1
    worksheet.write(start_row, start_col, gross_totals[i]) #3
    start_col = start_col + 1
    worksheet.write(start_row, start_col, net_totals[i]) #4
    start_col = start_col + 1
    worksheet.write(start_row, start_col, shipping_costs[i]) #5
    start_col = start_col + 1
    worksheet.write(start_row, start_col, shipping_promotions[i]) #6
    # worksheet.write(start_row, start_col, totals[i]) #7

    if(start_col == 6):
        start_col = 0
        start_row = start_row +1


#첫 줄에 데이터 입력
for n, info in enumerate(row_one_line):
    worksheet.write(0, n, info)



workbook.close()
driver.close()

