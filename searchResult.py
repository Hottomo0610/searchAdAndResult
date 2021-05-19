import time
import logging
from logging import getLogger, StreamHandler, Formatter
import re
from collections import defaultdict
import pprint
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
#from bs4 import BeautifulSoup
import json
import sys
import random
import os
import os.path
import json
import signal
import datetime
# from __future__ import print_function
import pickle
import gspread
from oauth2client.service_account import ServiceAccountCredentials


# 最低限の設定をしてchromedriverを起動させる関数
def start_chrome_driver():
    user_agent = ['Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.169 Safari/537.36',
              'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.121 Safari/537.36',
              'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.157 Safari/537.36',
              'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36',
              'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36'
              ]
    start = time.time()
    options = Options()
    options.add_argument('--headless') #headlessでchromeを起動
    options.add_argument("--disable-gpu")
    options.add_argument('--no-sandbox')
    options.add_argument('--lang=ja-JP')
    prefs = {"download.default_directory" : r"/vagrant/workspace/searchAdAndResult/"}
    options.add_experimental_option("prefs",prefs)
    options.add_argument('--user-agent=' + user_agent[random.randrange(0, len(user_agent), 1)])
    driver_path = "/usr/local/bin/chromedriver"
    driver = webdriver.Chrome(executable_path=driver_path, options=options)
    return driver

# テキストから<～>で挟まれた部分を削除する関数
def remove_bracket(text):
    check1 = True                                            # whileループの終了条件に使用
    check2 = True
    word_s = '<'                                            # タグの先頭<を検索する時に使用
    word_e = '>'                                            # タグの末尾>を検索する時に使用
    word_n = '\n'
 
    # 「<>」のセットが無くなるまでループ
    while check1 == True:
        start = text.find(word_s)                           # <が何番目の指標かを検索
        # もしword_sが無い場合はcheckをFにしてwhileを終了
        if start == -1:
            check1 = False
        # word_sが存在する場合
        else:
            end = text.find(word_e)                         # >が何番目の指標かを検索
            # もしword_eが無い場合はcheckをFにしてwhileを終了
            if end == -1:
                check1 = False
            # word_sとword_eの両方がセットである場合は<と>で囲まれた範囲を空白に置換（削除）する。
            else:
                remove_word = text[start:end + 1]           # 削除するワード（<と>で囲まれた所）をスライス
                text = text.replace(remove_word, '')        # remove_wordを空白に置換
    
    while check2 == True:
        start = text.find(word_n)                           # <が何番目の指標かを検索
        # もしword_sが無い場合はcheckをFにしてwhileを終了
        if start == -1:
            check2 = False
        # word_sが存在する場合
        else:
            end = text.find(word_e)                         # >が何番目の指標かを検索
            # もしword_eが無い場合はcheckをFにしてwhileを終了
            if end == -1:
                check2 = False
            # word_sとword_eの両方がセットである場合は<と>で囲まれた範囲を空白に置換（削除）する。
            else:
                remove_word = text[start:start + 1]           # 削除するワード（<と>で囲まれた所）をスライス
                text = text.replace(remove_word, '')        # remove_wordを空白に置換
    return text

pref_list = [
    "北海道",
    "青森県",
    "岩手県",
    "秋田県",
    "宮城県",
    "山形県",
    "福島県",
    "茨城県",
    "栃木県",
    "群馬県",
    "山梨県",
    "千葉県",
    "埼玉県",
    "東京都",
    "神奈川県",
    "新潟県",
    "長野県",
    "富山県",
    "石川県",
    "福井県",
    "静岡県",
    "愛知県",
    "岐阜県",
    "三重県",
    "滋賀県",
    "和歌山県",
    "奈良県",
    "京都府",
    "大阪府",
    "兵庫県",
    "鳥取県",
    "島根県",
    "岡山県",
    "広島県",
    "山口県",
    "香川県",
    "徳島県",
    "愛媛県",
    "高知県",
    "福岡県",
    "佐賀県",
    "長崎県",
    "宮崎県",
    "熊本県",
    "鹿児島県",
    "沖縄県"
]
pref_latitude = {
    "北海道": 43.06417,
    "青森県": 40.82444,
    "岩手県": 39.70361,
    "秋田県": 39.71861,
    "宮城県": 38.26889,
    "山形県": 38.24056,
    "福島県": 37.75,
    "茨城県": 36.34139,
    "栃木県": 36.56583,
    "群馬県": 36.39111,
    "山梨県": 35.66389,
    "千葉県": 35.60472,
    "埼玉県": 35.85694,
    "東京都": 35.68944,
    "神奈川県": 35.44778,
    "新潟県": 37.90222,
    "長野県": 36.65139,
    "富山県": 36.69528,
    "石川県": 36.59444,
    "福井県": 36.06528,
    "静岡県": 34.97694,
    "愛知県": 35.18028,
    "岐阜県": 35.39111,
    "三重県": 34.73028,
    "滋賀県": 35.00444,
    "和歌山県": 34.22611,
    "奈良県": 34.68528,
    "京都府": 35.02139,
    "大阪府": 34.68639,
    "兵庫県": 34.69139,
    "鳥取県": 35.50361,
    "島根県": 35.47222,
    "岡山県": 34.66167,
    "広島県": 34.39639,
    "山口県": 34.18583,
    "香川県": 34.34028,
    "徳島県": 34.06583,
    "愛媛県": 33.84167,
    "高知県": 33.55972,
    "福岡県": 33.60639,
    "佐賀県": 33.24944,
    "長崎県": 32.74472,
    "宮崎県": 31.91111,
    "熊本県": 32.78972,
    "鹿児島県": 31.56028,
    "沖縄県": 26.2125
}
pref_longitude = {
    "北海道": 141.34694,
    "青森県": 140.74,
    "岩手県": 141.1525,
    "秋田県": 140.1025,
    "宮城県": 140.87194,
    "山形県": 140.36333,
    "福島県": 140.46778,
    "茨城県": 140.44667,
    "栃木県": 139.88361,
    "群馬県": 139.06083,
    "山梨県": 138.56833,
    "千葉県": 140.12333,
    "埼玉県": 139.64889,
    "東京都": 139.69167,
    "神奈川県": 139.6425,
    "新潟県": 139.02361,
    "長野県": 138.18111,
    "富山県": 137.21139,
    "石川県": 136.62556,
    "福井県": 136.22194,
    "静岡県": 138.38306,
    "愛知県": 136.90667,
    "岐阜県": 136.72222,
    "三重県": 136.50861,
    "滋賀県": 135.86833,
    "和歌山県": 135.1675,
    "奈良県": 135.83278,
    "京都府": 135.75556,
    "大阪府": 135.52,
    "兵庫県": 135.18306,
    "鳥取県": 134.23833,
    "島根県": 133.05056,
    "岡山県": 133.935,
    "広島県": 132.45944,
    "山口県": 131.47139,
    "香川県": 134.04333,
    "徳島県": 134.55944,
    "愛媛県": 132.76611,
    "高知県": 133.53111,
    "福岡県": 130.41806,
    "佐賀県": 130.29889,
    "長崎県": 129.87361,
    "宮崎県": 131.42389,
    "熊本県": 130.74167,
    "鹿児島県": 130.55806,
    "沖縄県": 127.68111
}
fields = [
    "債務整理",
    "自己破産", 
    "闇金", 
    "刑事事件", 
    "残業代請求", 
    "不当解雇", 
    "遺産分割", 
    "遺留分", 
    "財産分与", 
    "親権", 
    "交通事故", 
    "養育費回収",
    "不動産投資詐欺", 
    "投資マンション被害",
    "誹謗中傷", 
    "削除請求", 
    "労働災害",
    "ビザ", 
    "永住許可", 
    "帰化申請", 
    "補助金申請",
    "アスベスト",
    "B型肝炎"
]

logger = getLogger("Logging")
logger.setLevel(logging.DEBUG)
stream_handler = StreamHandler()
stream_handler.setLevel(logging.DEBUG)
handler_format = Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
stream_handler.setFormatter(handler_format)
logger.addHandler(stream_handler)
fh = logging.FileHandler('/usr/local/searchAdAndResult/test.log')
logger.addHandler(fh)

driver = start_chrome_driver()

driver.execute_cdp_cmd(
    "Browser.grantPermissions",
    {
        "origin": "https://www.google.co.jp/",
        "permissions": ["geolocation"]
    },
)

nested_dict = {}
nested_dict = lambda: defaultdict(nested_dict)

results = nested_dict()

proc_name1 = "chrome"
proc_name2 = "chromedriver"

# i = 0

for index, field in enumerate(fields):
    search_result = []
    for name in pref_list:
        search_result_by_query = []
        
        word1 = field + " 弁護士 " + name
        if(field=="ビザ" or field=="永住許可" or field=="帰化申請" or field=="補助金申請"):
            word1 = field + " 代行 " + name
        
        word2 = field + " 相談 " + name
        querys = [word1, word2]
        
        if(field=="不動産投資詐欺" or field=="投資マンション被害"):
            querys.append(field)
        
        driver = start_chrome_driver()

        driver.execute_cdp_cmd(
            "Browser.grantPermissions",
            {
                "origin": "https://www.google.co.jp/",
                "permissions": ["geolocation"]
            },
        )
        
        driver.implicitly_wait(1)
        
        driver.execute_cdp_cmd(
            "Emulation.setGeolocationOverride",
            {
                "latitude": pref_latitude[name],
                "longitude": pref_longitude[name],
                "accuracy": 100,
            },
        )

        for word in querys:
            driver.get('https://www.google.co.jp')
    
            search_bar = driver.find_element_by_name("q")
            search_bar.send_keys(word)
            search_bar.submit()
    
            time.sleep(2)
    
            # 一番下までスクロール
            # driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            
            w = driver.execute_script('return document.body.scrollWidth')
            h = driver.execute_script('return document.body.scrollHeight')
            driver.set_window_size(w, h)
    
            # image_path = "/usr/local/searchAdAndResult/searchImage"+ str(i) +".jpg"
            # i= i+1
    
            # driver.get_screenshot_as_file(image_path)
    
            #リスティング広告の検索結果を取得    
            try:
                ads_div = driver.find_element_by_xpath("//*[@id='tads']")
                ads_tag_list = ads_div.find_elements_by_class_name("uEierd")
                # print(str(len(ads_tag_list)))
            except:
                ads_tag_list = []        
            try:    
                bottom_ads_div = driver.find_element_by_xpath("//*[@id='tadsb']")
                bottom_ads_tag_list = bottom_ads_div.find_elements_by_class_name("uEierd")
                # print(str(len(bottom_ads_tag_list)))
            except:
                bottom_ads_tag_list = []
            
            for j in range(1, len(ads_tag_list)+1): 
                ad_element_list = {}
    
                try:
                    atag = driver.find_element_by_xpath("//*[@id='tads']/div["+str(j)+"]/div/div/div[1]/a")
                    url = atag.get_attribute("href")
                    
                    title = title = driver.find_element_by_xpath("//*[@id='tads']/div["+str(j)+"]/div/div/div[1]/a/div[1]").text
                    
                    raw_detail = driver.find_element_by_xpath("//*[@id='tads']/div["+str(j)+"]/div/div/div[3]/div").text
                    detail = remove_bracket(raw_detail)
        
                    ad_element_list["title"] = title
                    ad_element_list["detail"] = detail
                    ad_element_list["url"] = url
                    ad_element_list["pref"] = name
                    ad_element_list["ad_or_organic"] = "リスティング広告"            
                    # logger.debug(word+" ad was success.")
                except:
                    logger.debug(word+"/ ads error occured")
                else:
                    search_result_by_query.append(ad_element_list)
    
                
            for k in range(1, len(bottom_ads_tag_list)+1):
                bottom_ad_element_list = {}            
                try:
                    url = driver.find_element_by_xpath("//*[@id='tadsb']/div["+str(k)+"]/div/div/div[1]/a").get_attribute("href")
                    # print(url)
                    title = driver.find_element_by_xpath("//*[@id='tadsb']/div["+str(k)+"]/div/div/div[1]/a/div[1]").text
                    # print(title)
                    raw_detail = driver.find_element_by_xpath("//*[@id='tadsb']/div["+str(k)+"]/div/div/div[3]/div").text
                    detail = remove_bracket(raw_detail)
                    # print(detail)               
                    bottom_ad_element_list["title"] = title
                    bottom_ad_element_list["detail"] = detail
                    bottom_ad_element_list["url"] = url
                    bottom_ad_element_list["pref"] = name
                    bottom_ad_element_list["ad_or_organic"] = "リスティング広告"
                    # logger.debug(word+" bottom ad was success.")              
                except:
                    logger.debug(word+"/ bottom ads error occured.")
                else:
                    search_result_by_query.append(bottom_ad_element_list)           
    
            #自然検索の検索結果を取得
            organic_divs = driver.find_element_by_id("rso")
            result_list = organic_divs.find_elements_by_class_name("g")
            # print(str(len(result_list)))
            # l = 1
    
            for l in range(1, len(result_list)+1):          
                organic_results_elements_list = {}
    
                # try:
                try:
                    url = driver.find_element_by_xpath("//*[@id='rso']/div["+str(l)+"]/div/div/div[1]/a").get_attribute("href")
                except:
                    try:
                        url_element = driver.find_element_by_xpath("//*[@id='rso']/div["+str(l)+"]/div/div/div/div[1]/a")
                        url = url_element.get_attribute("href")
                    except:
                        url = ""          
                try:
                    title = driver.find_element_by_xpath("//*[@id='rso']/div["+str(l)+"]/div/div/div[1]/a/h3").text
                except:
                    try:
                        title = driver.find_element_by_xpath("//*[@id='rso']/div["+str(l)+"]/div/div/div/div[1]/a/h3").text
                    except:
                        title = ""        
                try:
                    raw_detail_e = driver.find_element_by_xpath("//*[@id='rso']/div["+str(l)+"]/div/div/div[2]")
                    raw_detail = raw_detail_e.text
                except:
                    try:
                        raw_detail_e = driver.find_element_by_xpath("//*[@id='rso']/div["+str(l)+"]/div/div/div/div[2]")
                        raw_detail = raw_detail_e.text
                    except:
                        raw_detail = ""
                    
                detail = remove_bracket(raw_detail)
                
                organic_results_elements_list["title"] = title
                organic_results_elements_list["detail"] = detail
                organic_results_elements_list["url"] = url
                organic_results_elements_list["pref"] = name
                organic_results_elements_list["ad_or_organic"] = "自然検索結果"
    
                search_result_by_query.append(organic_results_elements_list)
                # logger.debug(word + " organic results is over.")
            
            results[field][word] = search_result_by_query
            
        driver.quit()

time.sleep(1)

api_scope = ['https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive']

sheet_id = "1Vh0O_rnwZO1B_sav4lAKOKWafYYsJe85kQRV8PK0iIg"

sheet_names = ["債務整理", 
              "自己破産", 
              "闇金", 
              "刑事事件", 
              "残業代請求", 
              "不当解雇", 
              "遺産分割", 
              "遺留分", 
              "財産分与", 
              "親権", 
              "交通事故", 
              "養育費回収",
              "不動産投資詐欺", 
              "投資マンション被害",
              "誹謗中傷", 
              "削除請求", 
              "労働災害",
              "ビザ", 
              "永住許可", 
              "帰化申請", 
              "補助金申請",
              "アスベスト",
              "B型肝炎"]

json_path = "/usr/local/searchAdAndResult/credential.json"

credentials_path = os.path.join(os.path.expanduser('~'), 'path', 'to', json_path)
credentials = ServiceAccountCredentials.from_json_keyfile_name(json_path, api_scope)

gc = gspread.authorize(credentials)

workbook = gc.open_by_key(sheet_id)

range = "A1:G9999"
for sheet_name in sheet_names:
    d = 0
    work_sheet = workbook.worksheet(sheet_name)
    work_sheet.clear()
    send_list = []
    attribute = ["年/月", "都道府県", "クエリ", "リスティング広告", "タイトル", "説明文", "リンク"]
    send_list.append(attribute)
    list1 = results[sheet_name]
    
    for query, list2 in list1.items():
        # limit = len(list2) + 2
        for i, search_result in enumerate(list2):
            if(d==0):
                today = datetime.datetime.now().strftime('%Y – %m – %d')
                send_list2 = [today]
                d = d + 1
            else:
                send_list2 = [""]
            
            if(i==0):
                send_list2.append(search_result['pref'])
                send_list2.append(query)
            else:
                send_list2.append("")
                send_list2.append("")
            
            send_list2.append(search_result['ad_or_organic'])
            send_list2.append(search_result['title'])
            send_list2.append(search_result['detail'])
            send_list2.append(search_result['url'])
            
            send_list.append(send_list2)
        
    work_sheet.update(range, send_list)

exit()
