import requests
import json
import pandas as pd
import time
from seleniumwire import webdriver # 需安裝：pip install selenium-wire
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
import re
import random
import zlib
from tqdm import tqdm
from bs4 import BeautifulSoup
from fake_useragent import UserAgent

keyword = '運動長褲'
page = 1
ecode = 'utf-8-sig'

ua = UserAgent()
# 從UserAgent物件中獲取一個隨機的User-Agent，用於發送HTTP請求，降低被網站阻擋或限制的風險
user_agent = ua.random
my_headers = {'user-agent': user_agent}
   
# 進入每個商品，抓取買家留言
def goods_comments(item_id, shop_id):
    url = 'https://shopee.tw/api/v4/item/get_ratings?filter=0&flag=1&itemid='+ str(item_id) + '&limit=50&offset=0&shopid=' + str(shop_id) + '&type=0'
    r = requests.get(url,headers = my_headers)
    st= r.text.replace("\\n","^n")
    st=st.replace("\\t","^t")
    st=st.replace("\\r","^r")
    
    gj=json.loads(st)
    return gj['data']['ratings']
    

# 進入每個商品，抓取賣家更細節的資料（商品文案、SKU）
# https://shopee.tw/api/v4/item/get?itemid=17652103038&shopid=36023817
def goods_detail(url, item_id, shop_id):
    # 2022/12/29 ivan，因shopee API新增了防爬蟲機制，header中多了「af-ac-enc-dat」參數，因解析不出此參數如何製成，只能使用土法煉鋼，一頁一頁進去攔封包
    driver.get(url) # 需要到那個頁面，才能度過防爬蟲機制
    time.sleep(random.randint(10,20))
    getPacket = ''
    for request in driver.requests:
        if request.response:
            # 挑出商品詳細資料的json封包
            if 'https://shopee.tw/api/v4/item/get?itemid=' + str(item_id) + '&shopid=' + str(shop_id) in request.url:
                # 此封包是有壓縮的，因此需要解壓縮
                try:
                    getPacket = zlib.decompress(
                        request.response.body,
                        16+zlib.MAX_WBITS
                        )
                    break
                except:
                    print('封包拆解有誤')
    if getPacket != '':
        gj=json.loads(getPacket)
        return gj['data']
    else:
        return getPacket

# 自動下載ChromeDriver
service = ChromeService(executable_path=ChromeDriverManager().install())

# 關閉通知提醒
options = webdriver.ChromeOptions()
prefs = {"profile.default_content_setting_values.notifications" : 2}
options.add_experimental_option("prefs",prefs)
# 不載入圖片，提升爬蟲速度
# options.add_argument('blink-settings=imagesEnabled=false') 

# 開啟瀏覽器
driver = webdriver.Chrome(service=service, chrome_options=options)
time.sleep(random.randint(5,10))

# 開啟網頁，進到首頁
driver.get('https://shopee.tw/buyer/login?next=https%3A%2F%2Fshopee.tw%2F')
time.sleep(random.randint(5,10))

# 自動輸入帳號和密碼
username = "***"  
password = "***"  

# HTML元素 id = "email"
username_input = driver.find_element("name", "loginKey")

# HTML元素 id = "pass"
password_input = driver.find_element("name", "password")

# 模擬使用者輸入帳密
username_input.send_keys(username)
password_input.send_keys(password)

# 模擬按下 Enter 鍵進行登入
password_input.send_keys(Keys.RETURN)
time.sleep(random.randint(5,10))

#---------- Part 1. 主要先抓下商品名稱與連結，之後再慢慢補上詳細資料 ----------
print('---------- 開始進行爬蟲 ----------')
tStart = time.time()#計時開始
# 準備用來存放資料的陣列
itemid = [] # 商品ID
shopid =[] # 賣家ID
name = [] # 商品名稱
link = [] # 商品連結
price = [] # 價格
for i in tqdm(range(int(page))):
    driver.get('https://shopee.tw/search?keyword=' + keyword + '&page=' + str(i))    
    time.sleep(random.randint(5,10))
    # 滾動頁面
    for scroll in range(0,7):
        driver.execute_script('window.scrollBy(0,1000)')
        time.sleep(5)
    
    # 2023/04/20 由於使用selenium取得商品有些不穩定，因此以下換成全部使用bs4去解析
    # 取得商品內容
    for block in driver.find_elements(by=By.XPATH, value='//*[@data-sqe="item"]'):
        # 將整個網站的Html進行解析
        soup = BeautifulSoup(block.get_attribute('innerHTML'), "html.parser").find('a')
        time.sleep(5)
        # 商品連結、商品ID、商家ID
        getID = soup.get('href')
        theitemid = int((getID[getID.rfind('.')+1:getID.rfind('?')]))
        theshopid = int(getID[ getID[:getID.rfind('.')].rfind('.')+1 :getID.rfind('.')]) 
        
        # 先整理標籤
        get_parent = soup.find('div',{"data-sqe":"name"}).parent.find_all("div", recursive=False)
        
        # 商品名稱
        if len(get_parent) >0: # 確認有資料再進行
            getname = get_parent[0].text
        else:
            print('抓不到資料，直接是空的')
            continue # 沒抓到這個商品就別爬了

        #價格
        if len(get_parent) >1: # 確認有資料再進行
            getSpan = get_parent[1].find_all('span')
            counter = []
            for j in getSpan:
                theprice = j.text
                theprice = theprice.replace('萬','')
                theprice = theprice.replace('$','')
                theprice = theprice.replace(',','')
                theprice = theprice.replace(' ','')

                if theprice != '':
                    counter.append(int(theprice))

            # 到這邊確認都有抓到資料，才將它塞入陣列，否則可能會有缺漏
            link.append("https://shopee.tw"+getID)
            itemid.append(theitemid)
            shopid.append(theshopid)
            name.append(getname)
            price.append(sum(counter)/len(counter))
        else:
            print('抓不到價格資料')
            continue # 沒抓到這個商品就別爬了

        break

    time.sleep(random.randint(20,30)) # 休息久一點

# 2023/04/20 先將每頁抓到的商品儲存下來，方便後續追蹤並爬蟲
dic = {
    '商品ID':itemid,
    '賣家ID':shopid,
    '商品名稱':name,
    '商品連結':link,
    '價格':price,
    '品牌': [ None for x in range(len(itemid)) ] ,
    '存貨數量':[ None for x in range(len(itemid)) ] ,
    '商品文案':[ None for x in range(len(itemid)) ] ,
    '上架時間':[ None for x in range(len(itemid)) ] ,
    '折數':[ None for x in range(len(itemid)) ] ,
    '可否搭配購買':[ None for x in range(len(itemid)) ] ,
    '可否大量批貨購買':[ None for x in range(len(itemid)) ] ,
    '選項':[ None for x in range(len(itemid)) ] ,
    '歷史銷售量':[ None for x in range(len(itemid)) ] ,
    '可否分期付款':[ None for x in range(len(itemid)) ] ,
    '是否官方賣家帳號':[ None for x in range(len(itemid)) ] ,
    '是否可預購':[ None for x in range(len(itemid)) ] ,
    '喜愛數量':[ None for x in range(len(itemid)) ] ,
    '商家地點':[ None for x in range(len(itemid)) ] ,
    'SKU':[ None for x in range(len(itemid)) ] ,
    '評價數量':[ None for x in range(len(itemid)) ] ,
    '五星':[ None for x in range(len(itemid)) ] ,
    '四星':[ None for x in range(len(itemid)) ] ,
    '三星':[ None for x in range(len(itemid)) ] ,
    '二星':[ None for x in range(len(itemid)) ] ,
    '一星':[ None for x in range(len(itemid)) ] ,
    '評分':[ None for x in range(len(itemid)) ] ,
    '資料已完整爬取':[ False for x in range(len(itemid)) ] ,
}
pd.DataFrame(dic).to_csv(f'C:\python-training\實作訓練\檔案存取區\{keyword}_商品資料.csv', encoding = ecode, index=False)


driver.close() 
