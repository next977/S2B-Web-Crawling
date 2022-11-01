# 처음 프로젝트 시작 시, 모듈 설치
# pycharm 사용시 Terimanal에서 명령어 입력
#pip install selenium
#pip install openpyxl
#pip install xlwt
#pip install requests
#pip install beautifulsoup4

# 크롬 브라우저를 띄우기 위해, 웹드라이버를 가져오기
from selenium import webdriver
from selenium.webdriver.chrome.options import Options


from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from openpyxl import Workbook
import requests
import xlwt
import urllib.request as urllib2
from bs4 import BeautifulSoup


print('1.Put Check Last Page *About:2050 (Ex:10')
page_number_Last = int(input('>>'))
print('####### Statring Cralwing ########')

workbook = xlwt.Workbook(encoding='utf-8')
workbook.default_style.font.heignt = 20 * 11

xlwt.add_palette_colour("lightgray", 0x21)
workbook.set_colour_RGB(0x21, 216, 216, 216)
xlwt.add_palette_colour("lightgreen", 0x22)
workbook.set_colour_RGB(0x22, 216, 228, 188)

worksheet = workbook.add_sheet('S2B WebCrawing',cell_overwrite_ok=True)
col_width_0 = 256 * 13
col_width_1 = 256 * 13
col_width_2 = 256 * 21
col_width_3 = 256 * 13
col_width_4 = 256 * 13
col_width_5 = 256 * 15
col_width_6 = 256 * 16
col_width_7 = 256 * 14
col_width_8 = 256 * 13
col_width_9 = 256 * 13
col_width_10 = 256 * 23

col_height_content = 48

worksheet.col(0).width = col_width_0
worksheet.col(1).width = col_width_1
worksheet.col(2).width = col_width_2
worksheet.col(3).width = col_width_3
worksheet.col(4).width = col_width_4
worksheet.col(5).width = col_width_5
worksheet.col(6).width = col_width_6
worksheet.col(7).width = col_width_7
worksheet.col(8).width = col_width_8
worksheet.col(9).width = col_width_9
worksheet.col(10).width = col_width_10

list_style = "font:height 180,bold on; pattern: pattern solid, fore_color lightgray; align: wrap on, vert centre, horiz center"
list_style_emphasize = "font:height 180, bold on;pattern: pattern solid, fore_color lightgreen; align:vert centre, horiz center"
# worksheet.write(0,0,"Date", xlwt.easyxf(list_style))
worksheet.write(0, 0, "No.", xlwt.easyxf(list_style))
worksheet.write(0, 1, "계약구분", xlwt.easyxf(list_style))
worksheet.write(0, 2, "거래구분", xlwt.easyxf(list_style))
worksheet.write(0, 3, "계약명", xlwt.easyxf(list_style_emphasize))
worksheet.write(0, 4, "수요기관", xlwt.easyxf(list_style))
worksheet.write(0, 5, "계약번호", xlwt.easyxf(list_style))
worksheet.write(0, 6, "계약상대자", xlwt.easyxf(list_style))
worksheet.write(0, 7, "대표자명", xlwt.easyxf(list_style))
worksheet.write(0, 8, "예정가격(원)", xlwt.easyxf(list_style))
worksheet.write(0, 9, "견적금액(원)", xlwt.easyxf(list_style_emphasize))
worksheet.write(0, 10, "결정일", xlwt.easyxf(list_style_emphasize))

row_marker = 0
column_marker = 0


#wb = Workbook()
#sheet = wb.active

page_results = []
search = "javascript:f_search();"
page_search_break = False

detail_page_datas = []
detail_page_datas_combine = []
data_2array = []
#data_2array.append([])
#data_2array[0].append('a')
#data_2array[0].append(20)
#data_2array.append([])
#data_2array[1].append('b')
#data_2array[1].append(30)

# search = str()+"일"
search_month = '월'
search_day = '일'
search_comma = ','
page_table_data_start = 1  # table 1부터 시작
page_table_data_stop = 10  # table 10이 마지막
page_number_start = 1

#page_number_Last = 2
row_marker = 0
column_marker = 0
excel_data_count = 0


url = "https://www.s2b.kr/S2BNCustomer/tcmo001.do?forwardName=list02"

# IP 우회 Tor Browser 이용
#chrome_options = Options()
#chrome_options.add_argument("--proxy-server=socks5://127.0.0.1:9150")
#driver = webdriver.Chrome(executable_path=r'.\chromedriver.exe', options=chrome_options)

# 크롬 드라이버로 크롬을 실행한다.
# chromedriver.exe 파일을 해당 프로젝트 안에 있어야함. ( 현재 폴더 경로 ./ )
driver = webdriver.Chrome('./chromedriver')


try:
    # 입찰정보 검색 페이지로 이동
    # Selenium
    # 3s 기다림
    driver.implicitly_wait(3)
    driver.get(url)

    for page_click_index in range(page_number_start, page_number_Last+1):
        driver.implicitly_wait(5)


        #page_table_data_parsing
        tr_list = driver.find_elements_by_tag_name('tbody')
        for tr in tr_list:
            #page_results.append(tr.text)
            a_tags = tr.find_elements_by_tag_name('a')
            if a_tags:
                for a_tag in a_tags:
                    href_tag = a_tag.get_attribute('href')
                    if href_tag.startswith("http") :
                        pass
                    else :
                        print(href_tag)
                        page_results.append(href_tag)
                        for page_search in page_results:
                            if search in page_search:
                                print("search Ok")
                                page_search_break = True
                                break
                if page_search_break == True:
                    break
            if page_search_break == True:
                break

        for i in range(page_table_data_start, page_table_data_stop+1) : #(1,11)
            #print("page_results[i]" ,page_results[i])
            page_str = page_results[i]
            str_split = page_str.split('\'')
            str_sum = '\"'+str_split[0]+'\''+str_split[1]+'\''+str_split[2]+'\''+str_split[3]+'\''+str_split[4]+'\"'
            print("str_sum",str_sum)

            driver.implicitly_wait(5)
            driver.find_element_by_xpath('//a[@href='+ str_sum +']').click()

            # detail page 모두 긁어서 리스트로 저장
            tr_list = driver.find_elements_by_tag_name('tr')

            search_break = False
            data_2array.append([])
            # detail page data parsing
            for tr in tr_list:
               #results.append(div.text) # results.append(div.text)
               table_list = tr.find_elements_by_tag_name('table')
               for table in table_list:
                   table_class = table.get_attribute('class')
                   if table_class == "td_dark_line":
                       #print("table.text = ",table.text)

                       #Detail Data parsing
                       detail_text = table.text
                       detail_text = detail_text.replace("계약구분", ':').replace("거래구분", ':').replace("계약명", ':').replace("수요기관", ':').replace("계약번호", ':').replace("계약상대자", ':').replace("대표자명", ':').replace("예정가격(원)", ':').replace("견적금액(원)", ':').replace("결정일", ':').replace('\n', ' ')
                       #print("detail_text=",detail_text)

                       #Detail_Data Array Add
                       detail_page_datas.append(detail_text)
                       print("detail_page_datas = ", detail_page_datas)

                       data_excel_str1 = []
                       data_excel_str2 = []
                       column_marker = 0
                       #for exit
                       for str_search in detail_page_datas:
                           #print("str_search = ",str_search)
                           if search_comma in str_search :
                               print("search Ok, detail_page_data_combine ")

                               #detail_page_data 합치기
                               data_2array[page_click_index-1].extend(detail_page_datas)

                               data_excel_str1 = detail_page_datas[0].split(':')
                               data_excel_str2 = detail_page_datas[1].split(':')
                               for excel_index in range(1,6): #1 ~ 5
                                   worksheet.write(row_marker + 1, column_marker + 1, data_excel_str1[excel_index])
                                   column_marker += 1 #열
                               for excel_index in range(1, 6): #1 ~ 5
                                   worksheet.write(row_marker + 1, column_marker + 1, data_excel_str2[excel_index])
                                   column_marker += 1 #열

                               excel_data_count += 1
                               row_marker += 1 #행
                               worksheet.write(row_marker, 0, excel_data_count)
                               search_break = True
                               break

                   if search_break == True :
                      break
               if search_break == True:
                  break

            #page 뒤로가기 클릭
            driver.back()
            driver.implicitly_wait(5)

            # detail page 마다의 data 비우기
            detail_page_datas.clear()

        #print("detail_data_combine", detail_page_datas_combine)
        print("data_2array[0] = ", data_2array[0])
        print("data_2array[1] = ", data_2array[1])
        print("data_2array[2] = ", data_2array[2])
        print("data_2array[3] = ", data_2array[3])
        print("data_2array[4] = ", data_2array[4])

        # Selenium page click
        page_str= "javascript:goList"
        page_current_dex = int(page_click_index)
        page_next_dex = int(page_click_index)+1
        print("page_current_dex", page_current_dex)
        print("page_next_dex",page_next_dex)
        page_next = page_str + '(' + str(page_next_dex) + ')'
        print("page_next ",page_next)
        # page_next = '\"'+ page + '\"'
        driver.find_element_by_xpath("//a[@href='" + page_next + "']").click()
        driver.implicitly_wait(5)

        # page_results data 비우기
        page_results.clear()
        print("page data parshing Ok")

except Exception as e:
    # 위 코드에서 에러가 발생한 경우 출력
    print(e)
finally:
    # 에러와 관계없이 실행되고, 크롬 드라이버를 종료
    driver.quit()
    workbook.save("S2B_Web_Crawling.xls")
    print("chrome close , excel save success")