import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from bs4 import BeautifulSoup
import xlsxwriter
import xlrd
from datetime import datetime, timedelta
import urllib.request

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
       
        url_hmall = "https://www.hyundaihmall.com/front/bmc/brodPordPbdv.do?cnt=0&date=" + search_day
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
        print(len(tds))

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
            print(handle)
            try:
                handle = urllib.request.urlopen(url_gs)
                print(handle)
            except:
                pass
        
        print("aaa")

        data = handle.read()
        soup = BeautifulSoup(data, 'html.parser', from_encoding='utf-8')

        articles = soup.findAll('article')

        for article in articles:
            time_table = soup.find('span', {'class':'times'})
            gs_time = time_table.text

            items = article.findAll('li', {'class':'prd-item'})
        
            gs_item_list = []
            for item in items:

                prd_dt = item.find('dt', {'class':'prd-name'})
                prd_name = prd_dt.find('a')

                if prd_name == None:
                    gs_item_list.append(prd_dt.text)
                else:
                    gs_item_list.append(prd_name.text.strip())
                    print(prd_name.text)

            one_day_list.append([today, '', gs_time , gs_item_list])

        homeshopping_list.append(one_day_list)

    return homeshopping_list

def crawling_ky_homeshopping(start_day, end_day):

    delta = end_day - start_day
    homeshopping_list = []

    for i in range(delta.days + 1):
        
        d = start_day + timedelta(days=i)
        today = d.strftime('%Y%m%d')
        search_day = d.strftime('%Y-%m-%d')

        one_day_list = []
       
        # KongYoung home shopping
        url_ky = "https://www.gongyoungshop.kr/tvshopping/selectScheduleSub.do?brcStdDate=20200610"
        print(url_ky)

        time.sleep(0.5)

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
        url_lotte = "http://www.lotteimall.com/main/searchTvPgmByDay.lotte?bd_date=20200609"
        print(url_lotte)

        time.sleep(0.5)

    return homeshopping_list

def crawling_nsshopping(start_day, end_day):

    delta = end_day - start_day
    homeshopping_list = []

    for i in range(delta.days + 1):
        
        d = start_day + timedelta(days=i)
        today = d.strftime('%Y%m%d')
        search_day = d.strftime('%Y-%m-%d')

        one_day_list = []
       
        # GS home shopping
        url_ns = "http://www.nsmall.com/TVHomeShoppingBrodcastingList?tab_gubun=1&tab_Week=1&tab_bord=0&selectDay=" + search_day + "&catalogId=18151&langId=-9&storeId=13001#goToLocation"
        print(url_gs)

        time.sleep(0.5)

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
        print(rdate)

        url = "http://display.cjmall.com/p/homeTab/main?hmtabMenuId=002409#bdDt%3A" + rdate
        browser.get(url)

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

def write_excel_file(result_list, view_all_item):

    workbook_name = "all_home_shopping_test.xlsx"
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

    worksheet_name ='CJ.oshopping'
    worksheet0 = workbook.add_worksheet(worksheet_name) 

    offset = 1

    worksheet0.write(0, 0, "날짜", filter_format3)
    worksheet0.write(0, 1, "분류/제목", filter_format3)
    worksheet0.write(0, 2, "시간", filter_format3)
    worksheet0.write(0, 3, "아이템", filter_format3)

    for d in range(len(result_list)):
        result_day = result_list[d]
        for i in range(len(result_day)):
            worksheet0.write(i+offset, 0, result_day[i][0], filter_format3)
            worksheet0.write(i+offset, 1, result_day[i][1], filter_format3)
            worksheet0.write(i+offset, 2, result_day[i][2], filter_format3)
            if view_all_item == 0:
                worksheet0.write(i+offset, 3, result_day[i][3][0], filter_format3)
            else:
                item_all = ''
                for j in range(len(result_day[i][3])):
                    if j == (len(result_day[i][3]) - 1):
                        item_all = item_all + result_day[i][3][j] 
                    else:
                        item_all = item_all + result_day[i][3][j] + '\n'
                    #worksheet0.write(i+j+offset, 3, result_day[i][3][j], filter_format3)
                worksheet0.write(i+offset, 3, item_all, filter_format3)
                #offset = offset + len(result_day[i][3])-1
        offset = offset + len(result_day)

    workbook.close()

def main():

    #[TODO] 
    ##티커머스...
    # K쇼핑, 신세계쇼핑, CJ오쇼핑플러스, 현대홈쇼핑플러스, 
    # 롯데원TV, GS마이샵, SK스토아, W쇼핑, 쇼핑엔티, 
    # NS홈쇼핑+, 홈앤쇼핑2채널, K쇼핑2채널

    # Options...
    start_day = datetime(2020,6,1)
    end_day = datetime(2020,6,2)
    #delta_days = end_day-start_day

    view_all_item = 1

    #input_file = "req_trade.xlsx"
    
    ## 라이브 홈쇼핑
    # CJ오쇼핑
    #result_list = crawling_cj_oshopping(start_day, end_day)

    # GS홈쇼핑
    #result_list = crawling_gs_homeshopping(start_day, end_day)

    #현대홈쇼핑
    #result_list = crawling_hyundai_shopping(start_day, end_day)

    #홈앤쇼핑
    result_list = crawling_home_and_shopping(start_day, end_day)

    #롯데홈쇼핑
    #result_list = crawling_lotte_homeshopping(start_day, end_day)

    #NS홈쇼핑
    #result_list = crawling_nsshooping(start_day, end_day)
    
    #공영홈쇼핑
    #result_list = crawling_ky_homeshopping(start_day, end_day)

    #print(result_list)

    write_excel_file(result_list, view_all_item)


# Main
if __name__ == "__main__":
    main()


