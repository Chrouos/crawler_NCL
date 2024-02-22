# from selenium import webdriver
from seleniumwire import webdriver  
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from time import sleep

import requests
from bs4 import BeautifulSoup
from time import sleep
import re
import pandas as pd
from fake_useragent import UserAgent

# - 正規化搜索指令
def generate_search_query(student_name=None, school_name=None, department_name=None, advisor_name=None):
    query_parts = []
    
    # 檢查並生成每個部分的搜索字符串
    if student_name is not None:
        query_parts.append(f'"{student_name}".au')
    if school_name is not None:
        query_parts.append(f'"{school_name}".asc')
    if department_name is not None:
        query_parts.append(f'"{department_name}".dp')
    if advisor_name is not None:
        query_parts.append(f'"{advisor_name}".ad')
    
    # 用' and '連接所有的搜索字符串部分
    search_query = ' and '.join(query_parts)
    return search_query

# - 上網
def surf(cookie, rs, r1=1, h1=1):
    res_get = rs.get(f'https://ndltd.ncl.edu.tw/cgi-bin/gs32/gsweb.cgi/ccd={cookie}/record?r1={r1}&h1={h1}')
    sleep(2)  

    return res_get

# - 重新獲取 cookies
def reload_cookies(url, query):
    
    ua = UserAgent()
    user_agent = ua.random
    
    chrome_opt = Options()
    chrome_opt.add_argument("start-maximized") # 開啟視窗最大化
    chrome_opt.add_argument("--incognito") # 設置隱身模式，可以避免個人化廣告，加速網頁瀏覽
    
    
    driver = webdriver.Chrome(options = chrome_opt)
    driver.delete_all_cookies() # 防止污染
    
    driver.get(url)

    sleep(2)
    driver.find_element(By.XPATH, '//a[@title="指令查詢"]').click()

    sleep(2)
    driver.find_element(By.ID, 'ysearchinput0').send_keys(f'{query}')

    sleep(0.5)
    driver.find_element(By.ID, 'gs32search').click()

    sleep(2)
    cookie = re.findall(r'ccd=(.*?)/', driver.current_url)[0]
    
    # 自動填入 headers
    headers = {}
    for request in driver.requests:
        if f"https://ndltd.ncl.edu.tw/cgi-bin/gs32/gsweb.cgi/ccd={cookie}/search" == str(request):
            headers = {k: v for k, v in request.headers.items()} 
            
    sleep(1)
    driver.quit()
    
    if headers == {}:
        headers = {
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7", 
            "Accept-Encoding": "gzip, deflate, br", 
            "Accept-Language": "zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7", 
            "Connection": "keep-alive",
            "Host": "ndltd.ncl.edu.tw",  #目標網站 
            "Origin": "https://ndltd.ncl.edu.tw",
            "Referer": "https://ndltd.ncl.edu.tw/cgi-bin/gs32/gsweb.cgi/ccd={}/search?mode=cmd".format(cookie),
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "same-origin",
            "Sec-Fetch-User": "?1",

            "User-Agent": user_agent,
            'Cookie': 'ccd={}'.format(cookie)
        }

    cookie, rs, res_post, h1  = re_post_request(cookie, query, headers, 0)
    return cookie, rs, res_post, headers, h1


# - 重新 Post Request
def re_post_request(cookie, query, headers, h1):
    
    payload = {     'qs0': query,
                    'qf0': '_hist_',
                    'gs32search.x': '27',
                    'gs32search.y': '9',
                    'displayonerecdisable': '1',
                    'dbcode': 'nclcdr',
                    'action':'',
                    'op':'',
                    'h':'',
                    'histlist':'1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13',
                    'opt': 'm',
                    '_status_': 'search__v2'}
    
    sleep(1)
    rs = requests.session()
        
    sleep(1)
    res_post = rs.post('https://ndltd.ncl.edu.tw/cgi-bin/gs32/gsweb.cgi/ccd={}/search'.format(cookie), data=payload, headers=headers)
    h1 += 1
    
    return cookie, rs, res_post, h1

# - 從學校名稱中提取學校名稱和系所名稱
def re_school_department(school_name):
    pattern = r'(?P<University>[^系]+大學)(?P<Department>[\u4e00-\u9fff]+)'
    # 使用正則表達式匹配學校名稱和系所名稱
    match = re.search(pattern, school_name)
    if match:
        university = match.group('University')
        department = match.group('Department')
    else:
        university = None
        department = None
        
    return university, department

