#-*- coding:utf-8 -*-
# Read urls from Excel file
# Parse information from itooza
# Write information to Excel file
import xlrd
import xlsxwriter
import os
from bs4 import BeautifulSoup
import urllib.request
import pickle
import getopt
import sys
from datetime import datetime, timedelta
import time

def main():

    config_start_year   = 2020
    config_start_month  = 6
    config_start_day    = 1
    config_end_year     = 2020
    config_end_month    = 6
    config_end_day      = 2
    
    start_day = datetime(config_start_year, config_start_month, config_start_day)
    end_day = datetime(config_end_year, config_end_month, config_end_day)
    delta = end_day - start_day

    gs_date_list = []
    gs_time_list = []
    gs_type_list = []
    gs_name_list = []
    gs_link_list = []

    hmall_date_list = []
    hmall_time_list = []
    hmall_type_list = []
    hmall_name_list = []
    hmall_link_list = []

    hs_date_list = []
    hs_time_list = []
    hs_type_list = []
    hs_name_list = []
    hs_link_list = []

    for i in range(delta.days + 1):
        
        d = start_day + timedelta(days=i)
        today = d.strftime('%Y%m%d')
        today2 = today[0:4] + '%2F' + today[4:6] + '%2F' + today[6:8]
        print(today2)
        #today = "20190610"
        lseq = "409904"
        # GS home shopping
        url_gs = "https://www.gsshop.com/shop/tv/tvScheduleDetail.gs?today=" + today + "&lseq=" + lseq
        print(url_gs)

        # CJ home shopping
        # https://display.cjmall.com/c/rest/tv/tvSchedule?bdDt=20190607&isMobile=false&broadType=live&isEmployee=false 
        url_cj = "https://display.cjmall.com/c/rest/tv/tvSchedule?bdDt=" + today + "&isMobile=false&broadType=live&isEmployee=false"

        # NS home shopping
        # https://www.nsmall.com/TVHomeShoppingBrodcastingList?tab_gubun=1&tab_Week=1&tab_bord=0&selectDay=2019-06-03&catalogId=18151&langId=-9&storeId=13001#goToLocation

        # GS 홈쇼핑
        handle = None
        while handle == None:
            try:
                handle = urllib.request.urlopen(url_gs)
                #print(handle)
            except:
                pass
                        
        data = handle.read()
        soup = BeautifulSoup(data, 'html.parser', from_encoding='utf-8')

        time_tables= soup.findAll('span', {'class':'times'})
        items = soup.findAll('ul')

        print(len(time_tables))
        print(len(items))

        for i in range(len(time_tables)):
            #prdinfo = items[i].findAll('dl', {'class':'prd-info'})
            prdname = items[i].findAll('dt', {'class':'prd-name'})
            #for j in range(len(prdname)):
            for j in range(1):
                labels = prdname[j].findAll('label')
                subitems = prdname[j].findAll('a')
                print("subitems", len(subitems), "labels", len(labels))
                if len(subitems) != 0:
                    gs_date_list.append(today)
                    if (len(labels) != 0):
                        gs_type_list.append(labels[0].text)
                    else:
                        gs_type_list.append("")
                    gs_name_list.append(subitems[0].text)
                    gs_link_list.append(subitems[0].attrs['href'])
                    gs_time_list.append(time_tables[i].text)
                
        # 현대 홈쇼핑
        # Hyundai home shopping
        # view-source:https://www.hyundaihmall.com/front/bmc/brodPordPbdv.do?cnt=0&date=20190612 
        url_hmall = "https://www.hyundaihmall.com/front/bmc/brodPordPbdv.do?cnt=0&date=" + today

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

        time_tables= soup.findAll('p', {'class':'time'})
        hosts= soup.findAll('span', {'class':'host'})
        items = soup.findAll('div', {'class':'prod_info'})
        
        print(len(time_tables))
        print(len(items))

        for i in range(len(time_tables)):
            prdname = items[i].findAll('p', {'class':'prod_tit'})
            labels = hosts[i].findAll('b')
            #for j in range(len(prdname)):
            #print(i)
            if len(prdname) != 0:
                for j in range(1):
                    subitems = prdname[j].findAll('a')
                    if len(subitems) != 0:
                        hmall_date_list.append(today)
                        hmall_type_list.append(labels[0].text)
                        hmall_name_list.append(subitems[0].text)
                        hmall_link_list.append(subitems[0].attrs['onclick'])
                        hmall_time_list.append(time_tables[i].text)

        # Home & Shopping
        # http://www.hnsmall.com/display/tvtable.do?from_date=2019%2F06%2F05&search_key= 
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

        tds= soup.findAll('td')
        print(len(tds))
        
        item_count = 0

        for i in range(len(tds)):
            if 'class' in tds[i].attrs:
                #print(tds[i].attrs['class'])
                classes = tds[i].attrs['class']
                if classes[0] == 'dateTime':
                    #print("dateTime")
                    time_tables= tds[i].findAll('span', {'class':'time'})
                    hosts= tds[i].findAll('strong', {'class':'tit'})
                    item_count = 0
                elif classes[0] == 'goods' and item_count == 0:
                    prdname = tds[i].findAll('div', {'class':'text'})
                    #print("prdname")
                    if len(prdname) != 0:
                        for j in range(1):
                            subitems = prdname[j].findAll('a')
                            #print(len(subitems))
                            if len(subitems) != 0:
                                hs_date_list.append(today)
                                hs_type_list.append(hosts[0].text)
                                hs_name_list.append(subitems[0].text)
                                hs_link_list.append(subitems[0].attrs['href'])
                                hs_time_list.append(time_tables[0].text)
                                item_count = 1


    ### PART II - Write information to new Excel file
    out_dir = time.strftime("%y%m%d_%H%M") 
    workbook_name = out_dir+"/HOME_SHOPPING_"+out_dir+".xlsx"
    
    cur_dir = os.getcwd()
    if not os.path.exists(out_dir):
        os.mkdir(out_dir)
    
    if os.path.isfile(os.path.join(cur_dir, workbook_name)):
        os.remove(os.path.join(cur_dir, workbook_name))
    workbook = xlsxwriter.Workbook(workbook_name)
    worksheet_1 = workbook.add_worksheet('gsshop')

    filter_format = workbook.add_format({'bold':True,
                                        'fg_color': '#D7E4BC'
                                        })
    filter_format2 = workbook.add_format({'bold':True
                                        })

    percent_format = workbook.add_format({'num_format': '0.00%'})

    roe_format = workbook.add_format({'bold':True,
                                      'underline': True,
                                      'num_format': '0.00%'})

    num_format = workbook.add_format({'num_format':'0.00'})
    num2_format = workbook.add_format({'num_format':'#,##0'})
    num3_format = workbook.add_format({'num_format':'#,##0.00',
                                      'fg_color':'#FCE4D6'})

    # Write filter
    #worksheet_1.set_column('A:A', 15)
    #worksheet_1.set_column('B:B', 15)
    #worksheet_1.set_column('C:C', 10)
    #worksheet_1.set_column('D:D', 30)
    worksheet_1.write(0, 0, "번호", filter_format)
    worksheet_1.write(0, 1, "날짜", filter_format)
    worksheet_1.write(0, 2, "방송시간", filter_format)
    worksheet_1.write(0, 3, "분류", filter_format)
    worksheet_1.write(0, 4, "아이템", filter_format)
    worksheet_1.write(0, 5, "링크", filter_format)

    for k in range(len(gs_name_list)):
        worksheet_1.write(1+k, 0, k+1)
        worksheet_1.write(1+k, 1, gs_date_list[k])
        worksheet_1.write(1+k, 2, gs_time_list[k])
        if gs_type_list[k].find("자막방송") == -1:
            worksheet_1.write(1+k, 3, gs_type_list[k].strip())
        else:
            words = gs_type_list[k].split()
            #print(words)
            worksheet_1.write(1+k, 3, words[1].replace(' ',''))
        worksheet_1.write(1+k, 4, gs_name_list[k])
        if gs_link_list[k].find("gsshop") == -1:
            worksheet_1.write(1+k, 5, "http://www.gsshop.com" + gs_link_list[k])
        else:
            worksheet_1.write(1+k, 5, gs_link_list[k])

    worksheet_2 = workbook.add_worksheet('hmall')

    worksheet_2.write(0, 0, "번호", filter_format)
    worksheet_2.write(0, 1, "날짜", filter_format)
    worksheet_2.write(0, 2, "방송시간", filter_format)
    worksheet_2.write(0, 3, "분류", filter_format)
    worksheet_2.write(0, 4, "아이템", filter_format)
    worksheet_2.write(0, 5, "링크", filter_format)

    for k in range(len(hmall_name_list)):
        worksheet_2.write(1+k, 0, k+1)
        worksheet_2.write(1+k, 1, hmall_date_list[k])
        worksheet_2.write(1+k, 2, hmall_time_list[k])
        worksheet_2.write(1+k, 3, hmall_type_list[k].strip())
        #if hmall_type_list[k].find("자막방송") == -1:
        #   worksheet_2.write(1+k, 3, hmall_type_list[k].strip())
        #else:
        #   words = hmall_type_list[k].split()
        #   #print(words)
        #   worksheet_2.write(1+k, 3, words[1].replace(' ',''))
        worksheet_2.write(1+k, 4, hmall_name_list[k])
        words = hmall_link_list[k].split('\'')
        worksheet_2.write(1+k, 5, "http://www.hyundaihmall.com" + words[1])
        #if hmall_link_list[k].find("gsshop") == -1:
        #   worksheet_2.write(1+k, 5, "http://www.gsshop.com" + hmall_link_list[k])
        #else:
        #   worksheet_2.write(1+k, 5, hmall_link_list[k])

    worksheet_3 = workbook.add_worksheet('home & shopping')

    worksheet_3.write(0, 0, "번호", filter_format)
    worksheet_3.write(0, 1, "날짜", filter_format)
    worksheet_3.write(0, 2, "방송시간", filter_format)
    worksheet_3.write(0, 3, "제목", filter_format)
    worksheet_3.write(0, 4, "세부내용", filter_format)
    worksheet_3.write(0, 5, "링크", filter_format)

    for k in range(len(hs_name_list)):
        worksheet_3.write(1+k, 0, k+1)
        worksheet_3.write(1+k, 1, hs_date_list[k])
        worksheet_3.write(1+k, 2, hs_time_list[k])
        worksheet_3.write(1+k, 3, hs_type_list[k].strip())
        #if hs_type_list[k].find("자막방송") == -1:
        #   worksheet_3.write(1+k, 3, hs_type_list[k].strip())
        #else:
        #   words = hs_type_list[k].split()
        #   #print(words)
        #   worksheet_3.write(1+k, 3, words[1].replace(' ',''))
        worksheet_3.write(1+k, 4, hs_name_list[k])
        #words = hs_link_list[k].split('\'')
        worksheet_3.write(1+k, 5, hs_link_list[k])
        #worksheet_3.write(1+k, 5, "http://www.hyundaihmall.com" + words[1])
        #if hs_link_list[k].find("gsshop") == -1:
        #   worksheet_3.write(1+k, 5, "http://www.gsshop.com" + hs_link_list[k])
        #else:
        #   worksheet_3.write(1+k, 5, hs_link_list[k])

    workbook.close()

# Main
if __name__ == "__main__":
    main()


