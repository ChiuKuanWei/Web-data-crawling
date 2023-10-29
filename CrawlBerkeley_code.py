import requests # 取得網頁回傳的資料
from bs4 import BeautifulSoup # 用來解析HTML結構的套件
from fake_useragent import UserAgent
import json
from datetime import datetime
import os

#region 爬取柏克萊書籍網站內容 搜尋:python 取得所需內容及商品對應圖片；取得區域:滑鼠右鍵->檢查->元素

# 創建UserAgent物件
ua = UserAgent()
# 從UserAgent物件中獲取一個隨機的User-Agent，用於發送HTTP請求，降低被網站阻擋或限制的風險
user_agent = ua.random
my_head = {'user-agent': user_agent}

url = f'https://search.books.com.tw/search/query/key/python/cat/all' 

page_response = requests.get(url, headers= my_head)
#print(page_response.text)

#由於 response.text 不是json 的字典格式, 所以不能用之前的json.loads 的功能
soup = BeautifulSoup(page_response.text , 'html.parser')
#print(soup.prettify())  #輸出排版後的HTML內容

current_datetime = datetime.now().strftime('%Y%m%d%H%M%S')

# 這個錯誤是由於在Windows中的文件路徑中有些字符被認為是無效的，導致無法正確建立或開啟檔案。為了避免這個問題，你可以使用os.path模塊來確保路徑是正確的。
# 同時，你也應該避免在檔案名稱中使用不允許的特殊字符
str_path = os.path.join(r'C:\python-training\實作訓練\檔案存取區', f'取得網頁元素_{current_datetime}.txt') # 預設路徑

if soup.prettify() is not None:  
    # 將格式化的字串寫入文字檔
    with open(str_path, 'w', encoding='utf-8') as txt_file:
        txt_file.write(soup.prettify())

    print(f'資料已儲存至文字檔：{str_path}')

# 抓取想要的內容:

# STEP.1 找大區域集合
boxs = soup.find_all('div',class_ = "table-td") 

# STEP.2 找大區域內所需的類別
box_a  = boxs[0].find('a',target="_blank") 
box_img = box_a.find('img',class_ = 'b-lazy') #b-lazy ,b-loaded 兩種屬性

# STEP.3 get-得到當下盒子裡面其他指定的特徵
img_link= box_img.get('data-src')
title   = box_a.get('title')
print('圖片連結:',img_link)
print('產品title',title)

# STEP.4 text-得到當下盒子內的紙條
price_box = boxs[0].find('ul',class_ = 'price clearfix')
p_list = price_box.find_all('b')
discount = p_list[0].text
price    = p_list[1].text
print(f'優惠價:{discount}折')
print(f'商品價格:{price}元')

# 使用 requests.get 下載圖片
response = requests.get(img_link)

if response.status_code == 200: #如果回傳OK
    # 取得圖片的二進位內容
    image_data = response.content

    # 儲存圖片
    image_filename = os.path.join(r'C:\python-training\實作訓練\檔案存取區', f'圖片_{current_datetime}.jpg')
    with open(image_filename, 'wb') as image_file:
        image_file.write(image_data)

    print(f'圖片已儲存至：{image_filename}')
else:
    print('無法下載圖片!')


#endregion