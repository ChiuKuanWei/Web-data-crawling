# 參考文獻 :「蝦皮爬蟲」最詳細手把手教學，商品資料＋留言評論－【附Python程式碼】

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

####################################################################################
###                       至此，請先停下來手動登入帳號，再往後執行                  ###
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
####################################################################################


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

# tEnd = time.time()#計時結束
# totalTime = int(tEnd - tStart)
# minute = totalTime // 60
# second = totalTime % 60
# print('資料儲存完成，花費時間（約）： ' + str(minute) + ' 分 ' + str(second) + '秒')



# #---------- Part 2. 補上商品的詳細資料，由於多設了爬取的標記，因此爬過的就不會再爬了 ----------
# print('---------- 開始進行爬蟲 ----------')
# tStart = time.time()#計時開始

# # 2023/04/20 先取得之前爬下來的紀錄
# getData = pd.read_csv(keyword +'_商品資料.csv')
# for i in tqdm(range(len(getData))):
#     data = []
#     # 2023/04/20 備標註已經抓過的，就不用再抓了，這樣就算之前爬到一半被中斷，也不會努力付諸東流
#     if getData.iloc[i]['資料已完整爬取']==True:
#         continue
#     data.append(getData.iloc[i]['商品ID'])
#     data.append(getData.iloc[i]['賣家ID'])
#     data.append(getData.iloc[i]['商品名稱'])
#     data.append(getData.iloc[i]['商品連結'])
#     data.append(getData.iloc[i]['價格'])
#     #請求商品詳細資料
#     itemDetail = goods_detail(url = data[3], item_id = data[0], shop_id = data[1])

#     # 抓不到資料就先跳過
#     if itemDetail == '':
#         print('抓不到商品詳細資料...\n')
#         continue

#     data.append(itemDetail['brand'])# 品牌
#     data.append(itemDetail['stock'])# 存貨數量
#     data.append(itemDetail['description'])# 商品文案
#     data.append(itemDetail['ctime'])# 上架時間
#     data.append(itemDetail['discount'])# 折數
#     data.append(itemDetail['can_use_bundle_deal'])# 可否搭配購買
#     data.append(itemDetail['can_use_wholesale'])# 可否大量批貨購買
#     data.append(itemDetail['tier_variations'])# 可否分期付款
#     data.append(itemDetail['historical_sold'])# 歷史銷售量
#     data.append(itemDetail['is_cc_installment_payment_eligible'])# 可否分期付款
#     data.append(itemDetail['is_official_shop'])# 是否官方賣家帳號
#     data.append(itemDetail['is_pre_order'])# 是否可預購
#     data.append(itemDetail['liked_count'])# 喜愛數量
#     data.append(itemDetail['shop_location'])# 商家地點
#     #SKU
#     all_sku=[]
#     for sk in itemDetail['models']:
#         all_sku.append(sk['name'])
#     data.append(all_sku)# SKU
#     data.append(itemDetail['cmt_count'])# 評價數量
#     data.append(itemDetail['item_rating']['rating_count'][5])# 五星
#     data.append(itemDetail['item_rating']['rating_count'][4])# 四星
#     data.append(itemDetail['item_rating']['rating_count'][3])# 三星
#     data.append(itemDetail['item_rating']['rating_count'][2])# 二星
#     data.append(itemDetail['item_rating']['rating_count'][1])# 一星
#     data.append(itemDetail['item_rating']['rating_star'])# 評分
#     data.append(True)# 資料已完整爬取

#     getData.iloc[i] = data #塞入所有資料
#     getData.to_csv(keyword +'_商品資料.csv', encoding = ecode, index=False)
    
#     time.sleep(random.randint(45,70)) # 休息久一點

#     # 每爬5個商品，會再有一次更長的休息
#     if i%5 == 0 :
#         time.sleep(random.randint(30,150)) 

# tEnd = time.time()#計時結束
# totalTime = int(tEnd - tStart)
# minute = totalTime // 60
# second = totalTime % 60
# print('資料儲存完成，花費時間（約）： ' + str(minute) + ' 分 ' + str(second) + '秒')



# #---------- Part 3. 補上留言資料 ----------
# print('---------- 開始進行爬蟲 ----------')
# tStart = time.time()#計時開始
# container_comment = pd.DataFrame()
# # 2023/05/16 先取得之前爬下來的紀錄
# getData = pd.read_csv(keyword +'_商品資料.csv')
# for i in tqdm(range(len(getData))):
    
#     # 2023/05/16 消費者評論詳細資料
#     theitemid = getData.iloc[i]['商品ID']
#     theshopid = getData.iloc[i]['賣家ID']
#     getname = getData.iloc[i]['商品名稱']
#     theprice = getData.iloc[i]['價格']
#     iteComment = goods_comments(item_id = theitemid, shop_id = theshopid)
    
#     # 2023/5/16，抓不到資料就先跳過
#     if iteComment == None:
#         continue

#     userid = [] #使用者ID
#     anonymous = [] #是否匿名
#     commentTime = [] #留言時間
#     is_hidden = [] #是否隱藏
#     orderid = [] #訂單編號
#     comment_rating_star = [] #給星
#     comment = [] #留言內容
#     product_SKU = [] #商品規格
    
#     for comm in iteComment:
#         userid.append(comm['userid'])
#         anonymous.append(comm['anonymous'])
#         commentTime.append(comm['ctime'])
#         is_hidden.append(comm['is_hidden'])
#         orderid.append(comm['orderid'])
#         comment_rating_star.append(comm['rating_star'])
#         try:
#             comment.append(comm['comment'])
#         except:
#             comment.append(None)
        
#         p=[]
#         for pro in comm['product_items']:
#             try:
#                 p.append(pro['model_name'])
#             except:
#                 p.append(None)
                
#         product_SKU.append(p)
        
#     commDic = {
#         '商品ID':[ theitemid for x in range(len(iteComment)) ],
#         '賣家ID':[ theshopid for x in range(len(iteComment)) ],
#         '商品名稱':[ getname for x in range(len(iteComment)) ],
#         '價格':[ int(theprice) for x in range(len(iteComment)) ],
#         '使用者ID':userid,
#         '是否匿名':anonymous,
#         '留言時間':commentTime,
#         '是否隱藏':is_hidden,
#         '訂單編號':orderid,
#         '給星':comment_rating_star,
#         '留言內容':comment,
#         '商品規格':product_SKU
#         }

#     container_comment = pd.concat([container_comment,pd.DataFrame(commDic)], axis=0)
#     container_comment.to_csv(keyword +'_留言資料.csv', encoding = ecode, index=False)

#     time.sleep(random.randint(35,60)) # 休息久一點

# tEnd = time.time()#計時結束
# totalTime = int(tEnd - tStart)
# minute = totalTime // 60
# second = totalTime % 60
# print('資料儲存完成，花費時間（約）： ' + str(minute) + ' 分 ' + str(second) + '秒')



driver.close() 
