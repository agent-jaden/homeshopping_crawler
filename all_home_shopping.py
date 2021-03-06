import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from bs4 import BeautifulSoup
import xlsxwriter
import xlrd
from datetime import datetime, timedelta
import urllib.request
import json
import re
import os

def crawling_hyundai_shopping(start_day, end_day):

    delta = end_day - start_day
    homeshopping_list = []

    today = datetime.now()

    for i in range(delta.days + 1):
        
        d = start_day + timedelta(days=i)
        search_day = d.strftime('%Y%m%d')

        del_today = today - d
        #print(del_today.days)
        #if del_today.days > 0:
        #    cal_cnt = -1 *((del_today.days +3) // 7)
        #else:
        #    cal_cnt = (del_today.days - 4) // 7 + 1
        cal_cnt = -1 *((del_today.days +3) // 7)
        #print(cal_cnt)
        
        one_day_list = []
       
        url_hmall = "http://www.hyundaihmall.com/front/bmc/brodPordPbdv.do?cnt=" + str(cal_cnt) + "&date=" + search_day
        print(url_hmall)

        handle = None
        while handle == None:
            try:
                handle = urllib.request.urlopen(url_hmall)
                #print(handle)
            except:
                pass
        
        data = handle.read()
        soup = BeautifulSoup(data, 'html.parser', from_encoding='utf-8')

        brod_prod = soup.find('div', {'class':'brod_prod_list'})
        lis = brod_prod.findAll('li')

        for li in lis:
            #print(li)
            time_table= li.find('p', {'class':'time'})
            title_span = li.find('span', {'class':'host'})
            title = title_span.find('b')

            hmall_item_list = []
            items = li.findAll('p', {'class':'prod_tit'})
            for item in items:

                item_name = item.find('a')
                hmall_item_list.append(item_name.text.strip())
        
            one_day_list.append([search_day, title.text, time_table.text, hmall_item_list])

        homeshopping_list.append(one_day_list)
    
    return homeshopping_list

def crawling_home_and_shopping(start_day, end_day):

    delta = end_day - start_day
    homeshopping_list = []
    
    for i in range(delta.days + 1):
        
        d = start_day + timedelta(days=i)
        today = d.strftime('%Y%m%d')
        today2 = today[0:4] + '%2F' + today[4:6] + '%2F' + today[6:8]

        one_day_list = []
       
        url_hs = "http://www.hnsmall.com/display/tvtable.do?from_date=" + today2 + "&search_key="
        print(url_hs)

        handle = None
        while handle == None:
            try:
                handle = urllib.request.urlopen(url_hs)
                #print(handle)
            except:
                pass
                        
        data = handle.read()
        soup = BeautifulSoup(data, 'html.parser', from_encoding='utf-8')

        tv_table = soup.find('table',{'class':'tvList'})
        tds= tv_table.findAll('td')
        #print(len(tds))

        first_time = 0

        for i in range(len(tds)):
            if 'class' in tds[i].attrs:
                #print(tds[i].attrs['class'])
                classes = tds[i].attrs['class']
                if classes[0] == 'dateTime':
                    if first_time == 1:
                        one_day_list.append([today, hosts[0].text, time_tables[0].text, hn_item_list])
                    else:
                        first_time = 1
                    #print("dateTime")
                    time_tables= tds[i].findAll('span', {'class':'time'})
                    hosts= tds[i].findAll('strong', {'class':'tit'})
                    #hn_item_list.clear()
                    hn_item_list = []
                elif classes[0] == 'goods':
                    prdname = tds[i].find('div', {'class':'text'})
                    #print(prdname)
                    if len(prdname) != 0:
                        subitems = prdname.find('a')
                        hn_item_list.append(subitems.text.strip().replace('\n',''))
                        #one_day_list.append([today, hosts[0].text, time_tables[0].text, hn_item_list])

        one_day_list.append([today, hosts[0].text, time_tables[0].text, hn_item_list])
            
        homeshopping_list.append(one_day_list)

    return homeshopping_list

# 날짜/분류(없음)/시간/아이템
def crawling_gs_homeshopping(start_day, end_day):

    delta = end_day - start_day
    homeshopping_list = []

    for i in range(delta.days + 1):
        
        d = start_day + timedelta(days=i)
        today = d.strftime('%Y%m%d')

        one_day_list = []
       
        #today = "20190610"
        lseq = "409904"
        # GS home shopping
        url_gs = "https://www.gsshop.com/shop/tv/tvScheduleDetail.gs?today=" + today + "&lseq=" + lseq
        print(url_gs)

        time.sleep(0.5)
        #handle = urllib.request.urlopen(url_gs)
        # GS 홈쇼핑
        handle = None
        while handle == None:
            #print(handle)
            try:
                handle = urllib.request.urlopen(url_gs)
                #print(handle)
            except:
                pass

        data = handle.read()
        soup = BeautifulSoup(data, 'html.parser', from_encoding='utf-8')

        articles = soup.findAll('article')

        for article in articles:
            time_table = article.find('span', {'class':'times'})
            gs_time = time_table.text
            #print(gs_time)

            items = article.findAll('li', {'class':'prd-item'})
        
            gs_item_list = []
            for item in items:

                prd_dt = item.find('dt', {'class':'prd-name'})
                prd_name = prd_dt.find('a')

                if prd_name == None:
                    gs_item_list.append(prd_dt.text)
                else:
                    gs_item_list.append(prd_name.text.strip())
                    #print(prd_name.text)

            one_day_list.append([today, '', gs_time , gs_item_list])

        homeshopping_list.append(one_day_list)

    return homeshopping_list

def crawling_ky_homeshopping(start_day, end_day):

    delta = end_day - start_day
    homeshopping_list = []

    for i in range(delta.days + 1):
        
        d = start_day + timedelta(days=i)
        today = d.strftime('%Y%m%d')
        #search_day = d.strftime('%Y-%m-%d')

        one_day_list = []
       
        # KongYoung home shopping
        url_ky = "https://www.gongyoungshop.kr/tvshopping/selectScheduleSub.do?brcStdDate=" + today
        print(url_ky)

        time.sleep(0.5)

        handle = None
        while handle == None:
            #print(handle)
            try:
                handle = urllib.request.urlopen(url_ky)
                #print(handle)
            except:
                pass
        
        data = handle.read()
        json_file = json.loads(data.decode('utf-8'))
        #print(json_file.keys())
        prod_list = json_file['prdList']
        #print(len(prod_list), prod_list[0].keys())

        ky_item_list = []
        prev_begin_time = ''
        prev_end_time = ''
        prev_title = ''
        for prod in prod_list:
            begin_time = prod['brcBgnDtm']
            end_time = prod['brcEndDtm']
            title = prod['brcPgmNm']
            item_name = prod['prdNm']
            if prev_begin_time == "" or prev_begin_time == begin_time:
                ky_item_list.append(item_name)
            else:
                one_day_list.append([today, prev_title, prev_begin_time + "~" + prev_end_time, ky_item_list])
                ky_item_list = []
                ky_item_list.append(item_name)

            prev_end_time = prod['brcEndDtm']
            prev_title = prod['brcPgmNm']
            prev_begin_time = prod['brcBgnDtm']

        one_day_list.append([today, title, begin_time + "~" + end_time, ky_item_list])
        homeshopping_list.append(one_day_list)

    return homeshopping_list

def crawling_lotte_homeshopping(start_day, end_day):

    delta = end_day - start_day
    homeshopping_list = []

    for i in range(delta.days + 1):
        
        d = start_day + timedelta(days=i)
        today = d.strftime('%Y%m%d')
        search_day = d.strftime('%Y-%m-%d')

        one_day_list = []
       
        # Lotte home shopping
        url_lotte = "http://www.lotteimall.com/main/searchTvPgmByDay.lotte?bd_date=" + today
        print(url_lotte)

        time.sleep(0.5)
        
        handle = None
        while handle == None:
            #print(handle)
            try:
                handle = urllib.request.urlopen(url_lotte)
                #print(handle)
            except:
                pass

        data = handle.read()
        soup = BeautifulSoup(data, 'html.parser', from_encoding='utf-8')

        tsitem_list = soup.find('div',{'class','rn_tsitem_list'})
        divs = tsitem_list.findAll('div', {'class','rn_tsitem_wrap'})

        for div in divs:
            time_table_div = div.find('div',{'class','rn_tsitem_caption'})
            time_table = time_table_div.find('span')
            #print(time_table.text)
            
            lotte_item_list = []
            first_item = div.find('a',{'class':'rn_tsitem_title'})
            if first_item != None:
                lotte_item_list.append(first_item.text.strip())

                items = div.findAll('a',{'class':'rn_more_item'})
                for item in items:
                    lotte_item_list.append(item.text.strip())

                one_day_list.append([today, '', time_table.text, lotte_item_list])

        homeshopping_list.append(one_day_list)

    return homeshopping_list

def crawling_nsshopping(start_day, end_day):

    delta = end_day - start_day
    homeshopping_list = []

    options = Options()
    options.headless = True
    #options.headless = False
    browser = webdriver.Chrome(executable_path="./chromedriver.exe", options=options)

    for i in range(delta.days + 1):
        
        d = start_day + timedelta(days=i)
        today = d.strftime('%Y%m%d')
        search_day = d.strftime('%Y-%m-%d')

        one_day_list = []
       
        # NS home shopping
        url_ns = "http://www.nsmall.com/TVHomeShoppingBrodcastingList?selectDay=" + search_day + "#goToLocation"
        browser.get(url_ns)
        print(url_ns)

        time.sleep(0.5)

        html = browser.page_source
        soup = BeautifulSoup(html, 'html.parser', from_encoding='utf-8')

        tv_table = soup.find('div',{'class':'tv_table mt40'})
        #print(tv_table)
        tds = tv_table.findAll('td')
        #print(len(tds))
        
        first_time = 0
        
        for i in range(len(tds)):
            #print(i, tds[i].text.strip())
            if 'class' in tds[i].attrs:
                #print(tds[i].attrs['class'])
                classes = tds[i].attrs['class']
                if classes[0] == 'air':
                    if first_time == 1:
                        one_day_list.append([today, host.text, time_table.text, ns_item_list])
                    else:
                        first_time = 1
                    #print("dateTime")
                    time_table= tds[i].find('em')
                    host= tds[i].find('strong', {'class':'txt'})
                    #hn_item_list.clear()
                    ns_item_list = []
                elif classes[0] == 'al':
                    prdname = tds[i].find('div', {'class':'item_info'})
                    #print(prdname)
                    if len(prdname) != 0:
                        subitems = prdname.findAll('a')
                        if len(subitems) == 2:
                            ns_item_list.append(subitems[0].text.strip().replace('\n','') + subitems[1].text.strip().replace('\n',''))
                        else:
                            ns_item_list.append(subitems[0].text.strip().replace('\n',''))

        one_day_list.append([today, host.text, time_table.text, ns_item_list])

        homeshopping_list.append(one_day_list)

    return homeshopping_list

# 날짜/분류/시간/아이템
def crawling_cj_oshopping(start_day, end_day):
    options = Options()
    options.headless = True
    #options.headless = False
    browser = webdriver.Chrome(executable_path="./chromedriver.exe", options=options)

    delta = end_day - start_day
    homeshopping_list = []

    for d in range(delta.days + 1):
        
        current_day = start_day + timedelta(days=d)
        rdate = current_day.strftime('%Y%m%d')
        #print(rdate)

        url = "http://display.cjmall.com/p/homeTab/main?hmtabMenuId=002409#bdDt%3A" + rdate
        browser.get(url)
        print(url)

        time.sleep(0.5)

        html = browser.page_source
        soup = BeautifulSoup(html, 'html.parser', from_encoding='utf-8')
        #print(soup)

        div = soup.find('div', {"class":"schedule_prod"})
        prods = div.find_all('ul', {"class":"list_schedule_prod"})
        states = div.find_all('div', {"class":"state_bar"})

        one_day_list = []
        #tvschedule_wrap
        for i in range(len(prods)):
            
            #print(prods[i].text, states[i].text)

            tv_time = states[i].find('span',{'class':'pgmDtm'})
            title = states[i].find('span',{'class':'txt_cate'})

            lis = prods[i].find_all('li')
            item_list = []
            for li in lis:
                a = li.find('a',{'class':'link_alaram'})
                #strings = list(a.strings)
                #print(a['data-item-nm'])
                item_list.append(a['data-item-nm'].strip())
                #print(strings)

            one_day_list.append([rdate, title.text, tv_time.text, item_list])
        
        homeshopping_list.append(one_day_list)
        time.sleep(0.5)

        #print(one_day_list)

    browser.close()

    return homeshopping_list

def write_excel_file(result_list, view_all_item, search_data, search_op):

    now = datetime.now()
    out_dir = now.strftime("%y%m%d_%H%M") 

    cur_dir = os.getcwd()
    if not os.path.exists(out_dir):
        os.mkdir(out_dir)

    workbook_name = out_dir+"/all_home_shopping_"+out_dir+".xlsx"
    workbook = xlsxwriter.Workbook(workbook_name)
    #print(result_list)

    filter_format = workbook.add_format({'bold':True, 'fg_color': '#D7E4BC'	})
    filter_format.set_border()
    filter_format2 = workbook.add_format({'bold':True })
    filter_format2.set_border()
    filter_format3 = workbook.add_format({})
    filter_format3.set_border()

    percent_format = workbook.add_format({'num_format': '0.00%'})
    num_format = workbook.add_format({'num_format':'0.00'})
    num_format.set_border()
    num2_format = workbook.add_format({'num_format':'#,##0'})
    num2_format.set_border()
    #num3_format = workbook.add_format({'num_format':'#,##0.00', 'fg_color':'#FCE4D6'})

    if search_op == 1:
        #search_list = []
        for i in range(len(search_data)):
            search_sublist = []
            search_data[i].append(search_sublist)


    for n in range(len(result_list)):

        result_sheet = result_list[n][1]

        worksheet_name = result_list[n][0]
        worksheet0 = workbook.add_worksheet(worksheet_name) 
        print(worksheet_name)

        offset = 1

        worksheet0.set_column(0,0,20)
        worksheet0.set_column(0,1,20)
        worksheet0.set_column(0,2,20)

  
        worksheet0.write(0, 0, "날짜", filter_format3)
        worksheet0.write(0, 1, "분류/제목", filter_format3)
        worksheet0.write(0, 2, "시간", filter_format3)
        worksheet0.write(0, 3, "아이템", filter_format3)

        for d in range(len(result_sheet)):
            result_day = result_sheet[d]
            #print(result_day)
            #print(result_day[0][0])
            for i in range(len(result_day)):
                worksheet0.write(i+offset, 0, result_day[i][0], filter_format3)
                worksheet0.write(i+offset, 1, result_day[i][1], filter_format3)
                worksheet0.write(i+offset, 2, result_day[i][2], filter_format3)
                if view_all_item == 0:
                    worksheet0.write(i+offset, 3, result_day[i][3][0], filter_format3)
                else:
                    item_all = ''
                    for j in range(len(result_day[i][3])):
                        if result_day[i][3][j] != None:
                            if j == (len(result_day[i][3]) - 1):
                                item_all = item_all + result_day[i][3][j] 
                            else:
                                item_all = item_all + result_day[i][3][j] + '\n'
                        #worksheet0.write(i+j+offset, 3, result_day[i][3][j], filter_format3)
                    worksheet0.write(i+offset, 3, item_all, filter_format3)
                    #offset = offset + len(result_day[i][3])-1

                if search_op == 1:
                    for s in range(len(search_data)):
                        # Doing search
                        #print(search_data[s][1], result_list[n][0])
                        if search_data[s][1] == result_list[n][0]:
                            #print(search_data[s][2])
                            if re.compile(search_data[s][2]).search(item_all):
                                #print("search", item_all)
                                search_data[s][3].append([result_day[i][0], result_day[i][1],result_day[i][2], item_all])

            offset = offset + len(result_day)

    #print(search_data)

    if search_op == 1:
        worksheet0 = workbook.add_worksheet("Search_result")
        offset = 1

        worksheet0.write(0, 0, "기업", filter_format3)
        worksheet0.write(0, 1, "홈쇼핑", filter_format3)
        worksheet0.write(0, 2, "검색 아이템", filter_format3)
        worksheet0.write(0, 3, "날짜", filter_format3)
        worksheet0.write(0, 4, "분류/제목", filter_format3)
        worksheet0.write(0, 5, "시간", filter_format3)
        worksheet0.write(0, 6, "아이템", filter_format3)

        for i in range(len(search_data)):
            print(len(search_data[i][3]))
            if len(search_data[i][3]) == 0:
                worksheet0.write(offset, 0, search_data[i][0], filter_format3)
                worksheet0.write(offset, 1, search_data[i][1], filter_format3)
                worksheet0.write(offset, 2, search_data[i][2], filter_format3)
                worksheet0.write(offset+j, 3, '', filter_format3)
                worksheet0.write(offset+j, 4, '', filter_format3)
                worksheet0.write(offset+j, 5, '', filter_format3)
                worksheet0.write(offset+j, 6, '', filter_format3)
                offset = offset + 1
            else:
                for j in range(len(search_data[i][3])):
                    worksheet0.write(offset+j, 0, search_data[i][0], filter_format3)
                    worksheet0.write(offset+j, 1, search_data[i][1], filter_format3)
                    worksheet0.write(offset+j, 2, search_data[i][2], filter_format3)
                    worksheet0.write(offset+j, 3, search_data[i][3][j][0], filter_format3)
                    worksheet0.write(offset+j, 4, search_data[i][3][j][1], filter_format3)
                    worksheet0.write(offset+j, 5, search_data[i][3][j][2], filter_format3)
                    worksheet0.write(offset+j, 6, search_data[i][3][j][3], filter_format3)
                offset = offset + j+1

        worksheet0 = workbook.add_worksheet("Search_Count")
        offset = 1
        
        worksheet0.write(0, 0, "기업", filter_format3)
        worksheet0.write(0, 1, "홈쇼핑", filter_format3)
        worksheet0.write(0, 2, "검색 아이템", filter_format3)
        worksheet0.write(0, 3, "검색 횟수", filter_format3)

        for i in range(len(search_data)):
            worksheet0.write(offset, 0, search_data[i][0], filter_format3)
            worksheet0.write(offset, 1, search_data[i][1], filter_format3)
            worksheet0.write(offset, 2, search_data[i][2], filter_format3)
            worksheet0.write(offset, 3, len(search_data[i][3]), filter_format3)
            offset = offset + 1
    
    workbook.close()

def read_excel_file(input_file):

    read_data = []

    workbook_name = input_file
    workbook = xlrd.open_workbook(workbook_name)
    sheet_list = workbook.sheets()
    sheet1 = sheet_list[0]

    num_req = int(sheet1.cell(0,0).value)

    for i in range(num_req):
        corp_name = sheet1.cell(i+2,1).value
        hs_name = sheet1.cell(i+2,2).value
        search_item = sheet1.cell(i+2,3).value

        print("Read Data : ", corp_name, hs_name, search_item)
        read_data.append([corp_name, hs_name, search_item])
       
    return read_data

def scrape_homeshopping(input_file):

    #[TODO] 
    ##티커머스...
    # K쇼핑, 신세계쇼핑, CJ오쇼핑플러스, 현대홈쇼핑플러스, 
    # 롯데원TV, GS마이샵, SK스토아, W쇼핑, 쇼핑엔티, 
    # NS홈쇼핑+, 홈앤쇼핑2채널, K쇼핑2채널

    # Options...
    start_day = datetime(2020,11,1)
    end_day = datetime(2020,11,30)
    #delta_days = end_day-start_day
    select_cj       = 1
    select_gs       = 1
    select_hs       = 1
    select_hn       = 1
    select_lotte    = 1
    select_ns       = 1
    select_ky       = 1

    search_op       = 1

    view_all_item = 1

    result_list = []
    
    ## 라이브 홈쇼핑
    # CJ오쇼핑
    if select_cj == 1:
        print("CJ오쇼핑")
        result_sub_list = crawling_cj_oshopping(start_day, end_day)
        result_list.append(["CJ오쇼핑", result_sub_list])

    # GS홈쇼핑
    if select_gs == 1:
        print("GS홈쇼핑")
        result_sub_list = crawling_gs_homeshopping(start_day, end_day)
        result_list.append(["GS홈쇼핑", result_sub_list])

    #현대홈쇼핑
    if select_hs == 1:
        print("현대홈쇼핑")
        result_sub_list = crawling_hyundai_shopping(start_day, end_day)
        result_list.append(["현대홈쇼핑", result_sub_list])

    #홈앤쇼핑
    if select_hn == 1:
        print("홈앤쇼핑")
        result_sub_list = crawling_home_and_shopping(start_day, end_day)
        result_list.append(["홈앤쇼핑", result_sub_list])

    #롯데홈쇼핑
    if select_lotte == 1:
        print("롯데홈쇼핑")
        result_sub_list = crawling_lotte_homeshopping(start_day, end_day)
        result_list.append(["롯데홈쇼핑", result_sub_list])

    #NS홈쇼핑
    if select_ns == 1:
        print("NS홈쇼핑")
        result_sub_list = crawling_nsshopping(start_day, end_day)
        result_list.append(["NS홈쇼핑", result_sub_list])
    
    #공영홈쇼핑
    if select_ky == 1:
        print("공영홈쇼핑")
        result_sub_list = crawling_ky_homeshopping(start_day, end_day)
        result_list.append(["공영홈쇼핑", result_sub_list])

    #print(result_list)
    
    
    search_data = read_excel_file(input_file)

    write_excel_file(result_list, view_all_item, search_data, search_op)

def find_homeshopping(input_file):

    # Make search result

    # RAW data fiils
    data_files =    ["2020/all_home_shopping_2001.xlsx"]
    '''
    data_files =    [
                    "2020/all_home_shopping_2001.xlsx",
                    "2020/all_home_shopping_2002.xlsx",
                    "2020/all_home_shopping_2003.xlsx",
                    "2020/all_home_shopping_2004.xlsx",
                    "2020/all_home_shopping_2005.xlsx",
                    "2020/all_home_shopping_2006.xlsx",
                    "2020/all_home_shopping_2007.xlsx",
                    "2020/all_home_shopping_2008.xlsx",
                    "2020/all_home_shopping_2010.xlsx",
                    "2020/all_home_shopping_2011.xlsx"
                    ]
                    '''

    # Read result list
    # Homeshopping list -> one day list
    # search day / title / time_table  / item

    homeshopping_file_list = []

    read_files_list = []

    for file_name in data_files:

        read_file = []
        
        print("File name : ", file_name)
        workbook_name = file_name
        workbook = xlrd.open_workbook(workbook_name)
        sheet_list = workbook.sheets()

        for sheet1 in sheet_list[0:7]:

            read_homeshopping = []
            print(sheet1.nrows, sheet1.name)
            for i in range(sheet1.nrows-1):

                date    =   sheet1.cell(i+1,0).value
                title   =   sheet1.cell(i+1,1).value
                time    =   sheet1.cell(i+1,2).value
                item    =   sheet1.cell(i+1,3).value

                read_homeshopping.append([date, title, time, item])

            read_file.append([sheet1.name, read_homeshopping])
            print(len(read_file))
            #print(read_homeshopping)

        read_files_list.append(read_file)


    # corp_name, hs_name, search_item
    search_data = read_excel_file(input_file)

    for i in range(len(search_data)):
        search_sublist = []
        search_data[i].append(search_sublist)

    for result_file in read_files_list:
        for result_list in result_file:
            print(len(result_list[1]))
            for n in range(len(result_list[1])):
                #print(n)
                
                homeshopping_name = result_list[0]

                date    =   result_list[1][n][0]
                title   =   result_list[1][n][1]
                time    =   result_list[1][n][2]
                item_all=   result_list[1][n][3]

                for s in range(len(search_data)):
                    # Doing search
                    #print(search_data[s][1], result_list[n][0])
                    if search_data[s][1] == homeshopping_name:
                        #print(search_data[s][2])
                        if re.compile(search_data[s][2]).search(item_all):
                            print("search : ", date, title, time)
                            search_data[s][3].append([date, title, time, item_all])

    # Write search result
    now = datetime.now()
    out_dir = now.strftime("%y%m%d_%H%M") 

    cur_dir = os.getcwd()
    if not os.path.exists(out_dir):
        os.mkdir(out_dir)

    workbook_name = out_dir+"/search_result_"+out_dir+".xlsx"
    workbook = xlsxwriter.Workbook(workbook_name)

    filter_format = workbook.add_format({'bold':True, 'fg_color': '#D7E4BC'	})
    filter_format.set_border()
    filter_format2 = workbook.add_format({'bold':True })
    filter_format2.set_border()
    filter_format3 = workbook.add_format({})
    filter_format3.set_border()

    percent_format = workbook.add_format({'num_format': '0.00%'})
    num_format = workbook.add_format({'num_format':'0.00'})
    num_format.set_border()
    num2_format = workbook.add_format({'num_format':'#,##0'})
    num2_format.set_border()

    worksheet0 = workbook.add_worksheet("Search_result")
        
    offset = 1

    worksheet0.write(0, 0, "기업", filter_format3)
    worksheet0.write(0, 1, "홈쇼핑", filter_format3)
    worksheet0.write(0, 2, "검색 아이템", filter_format3)
    worksheet0.write(0, 3, "날짜", filter_format3)
    worksheet0.write(0, 4, "분류/제목", filter_format3)
    worksheet0.write(0, 5, "시간", filter_format3)
    worksheet0.write(0, 6, "아이템", filter_format3)

    for i in range(len(search_data)):
        print(len(search_data[i][3]))
        if len(search_data[i][3]) == 0:
            worksheet0.write(offset, 0, search_data[i][0], filter_format3)
            worksheet0.write(offset, 1, search_data[i][1], filter_format3)
            worksheet0.write(offset, 2, search_data[i][2], filter_format3)
            worksheet0.write(offset+j, 3, '', filter_format3)
            worksheet0.write(offset+j, 4, '', filter_format3)
            worksheet0.write(offset+j, 5, '', filter_format3)
            worksheet0.write(offset+j, 6, '', filter_format3)
            offset = offset + 1
        else:
            for j in range(len(search_data[i][3])):
                worksheet0.write(offset+j, 0, search_data[i][0], filter_format3)
                worksheet0.write(offset+j, 1, search_data[i][1], filter_format3)
                worksheet0.write(offset+j, 2, search_data[i][2], filter_format3)
                worksheet0.write(offset+j, 3, search_data[i][3][j][0], filter_format3)
                worksheet0.write(offset+j, 4, search_data[i][3][j][1], filter_format3)
                worksheet0.write(offset+j, 5, search_data[i][3][j][2], filter_format3)
                worksheet0.write(offset+j, 6, search_data[i][3][j][3], filter_format3)
            offset = offset + j+1

    worksheet0 = workbook.add_worksheet("Search_Count")
    offset = 1
    
    worksheet0.write(0, 0, "기업", filter_format3)
    worksheet0.write(0, 1, "홈쇼핑", filter_format3)
    worksheet0.write(0, 2, "검색 아이템", filter_format3)
    worksheet0.write(0, 3, "검색 횟수", filter_format3)

    for i in range(len(search_data)):
        worksheet0.write(offset, 0, search_data[i][0], filter_format3)
        worksheet0.write(offset, 1, search_data[i][1], filter_format3)
        worksheet0.write(offset, 2, search_data[i][2], filter_format3)
        worksheet0.write(offset, 3, len(search_data[i][3]), filter_format3)
        offset = offset + 1
    
    workbook.close()



def main():

    input_file = "homeshopping_checklist.xlsx"
    
    #scrape_homeshopping(input_file)

    find_homeshopping(input_file)

    


# Main
if __name__ == "__main__":
    main()


