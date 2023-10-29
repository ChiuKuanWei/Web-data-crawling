from tkinter import simpledialog, filedialog, Tk
import pandas as pd
import numpy as np
import requests
import json
from fake_useragent import UserAgent
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

#region Json回傳內容

# {
# {
#   "QTime": 48,
#   "totalRows": 7392,
#   "totalPage": 100,
#   "range": {
#     "min": "",
#     "max": ""
#   },
#   "cateName": "",
#   "q": "iphone 14 pro",
#   "subq": "",
#   "token": [
#     "iphone",
#     "14",
#     "pro"
#   ],
#   "isMust": 1,
#   "prods": [
#     {
#       "Id": "DYAJ7F-A900FH2OW",
#       "cateId": "DYAJ7F",
#       "picS": "/items/DYAJ7FA900FH2OW/000001_1662641863.jpg",
#       "picB": "/items/DYAJ7FA900FH2OW/000001_1662641863.jpg",
#       "name": "Apple iPhone 14 Pro (128G)-銀色(MQ023TA/A)",
#       "describe": "銀色★送無線充電板Apple iPhone 14 Pro (128G)-銀色(MQ023TA/A)\\r\\npchome安心保\r\n• pchome安心保 powered by applecare services\r\n需另外選購，購買 apple 商品後即可立即投保，\r\n請點此查詢價
# 格及保障內容。\r\n• 6.1 吋超 retina xdr 顯示器，具備永遠顯示與 promotion 自動適應更新頻率技術\r\n• 動態島功能，靈活多變的 iphone 新玩法\r\n• 4800 萬像素主相機，解析度最高可達 4 倍\r\n• 電影級模式，現在支援最
# 高可達 30 fps 的 4k 杜比視界\r\n• 動作模式，拍攝手持影片又順又穩\r\n• 重大安全技術—車禍偵測，關鍵時刻為你撥打求救電話\r\n• 滿足一天所需的電池續航力，影片播放時間最長可達 23 小時\r\n• a16 仿生，極致的智慧型手
# 機晶片。超飆速 5g 行動網路\r\n• 採用超瓷晶盾並且防潑抗水，耐用性皆為業界領先\r\n• ios 16 帶來更多個性化設定、溝通與共享的新方式\r\n• ncc許可字號：ccai225g0090t9\r\n\r\n\r\n\r\n↑詳情請點圖片了解\r\n\r\n舊機不
# 擔心，回收來pc\r\n↑舊愛不去新歡不來\r\n\r\n法律聲明\r\n1.「sos 緊急服務」功能須透過行動網路連線或 wi-fi 通話使用。\r\n2.顯示器採用圓角設計。以矩形量測時，螢幕的對角線長度為 6.12 吋。實際可視區較小。\r\n3.電 
# 池使用時間視使用情況及配置狀態而異。請參閱 apple.com tw batteries，以取得進一步資訊。\r\n4.須使用數據方案。5g 僅於特定地區並透過特定電信業者提供。連線速度會因所在地情況及電信業者而異。如需 5g 支援的詳細資訊 
# ，請洽詢你的電信業者並參閱 apple.com tw iphone cellular。\r\n5.iphone 14 pro 具備防潑抗水與防塵功能，並且已在受控管的實驗室環境條件下測試，達到 iec 60529 標準的 ip68 等級 (在最深達 6 公尺水中最長可達 30 分鐘
# )。防潑抗水與防塵功能非永久狀態，並可能隨日常使用造成的耗損而下降。請勿為潮濕的 iphone 充電；請參閱使用手冊了解清潔與乾燥處理說明。液體造成的損壞並不在保固範圍內。\r\n6.部分功能可能未在所有國家或地區提供。\r\n※相關注意事項，請見下方備註欄。",
#       "price": 32388,
#       "originPrice": 32388,
#       "author": "",
#       "brand": "",
#       "publishDate": "",
#       "isPChome": 1,
#       "isNC17": 0,
#       "couponActid": [],
#       "BU": "ec"
#     },
#     {
#       "Id": "DYAJ88-A900FHDI4",
#       "cateId": "DYAJ88",
#       "picS": "/items/DYAJ88A900FHDI4/000001_1686621881.jpg",
#       "picB": "/items/DYAJ88A900FHDI4/000001_1686621881.jpg",
#       "name": "Apple iPhone 14 Pro (256G)-銀色(MQ103TA/A)",
#       "describe": "銀色★送無線充電板Apple iPhone 14 Pro (256G)-銀色(MQ103TA/A)\\r\\n市價$38400．特價↘$３６３８８\r\n現買現省$2012\r\n限量特殺∼隨時回價∼！！\r\npchome安心保\r\n• pchome安心保 powered by applecare services\r\n需另外選購，購買 apple 商品後即可立即投保，\r\n請點此查詢價格及保障內容。\r\n• 6.1 吋超 retina xdr 顯示器，具備永遠顯示與 promotion 自動適應更新頻率技術\r\n• 動態島功能，靈活多變的 iphone 新 
# 玩法\r\n• 4800 萬像素主相機，解析度最高可達 4 倍\r\n• 電影級模式，現在支援最高可達 30 fps 的 4k 杜比視界\r\n• 動作模式，拍攝手持影片又順又穩\r\n• 重大安全技術—車禍偵測，關鍵時刻為你撥打求救電話\r\n• 滿足一天
# 所需的電池續航力，影片播放時間最長可達 23 小時\r\n• a16 仿生，極致的智慧型手機晶片。超飆速 5g 行動網路\r\n• 採用超瓷晶盾並且防潑抗水，耐用性皆為業界領先\r\n• ios 16 帶來更多個性化設定、溝通與共享的新方式\r\n• ncc許可字號：ccai225g0090t9\r\n\r\n\r\n\r\n↑詳情請點圖片了解\r\n\r\n↑詳情請點圖片了解\r\n\r\n舊機不擔心，回收來pc\r\n↑舊愛不去新歡不來\r\n\r\n法律聲明\r\n1.「sos 緊急服務」功能須透過行動網路連線或 wi-fi 通
# 話使用。\r\n2.顯示器採用圓角設計。以矩形量測時，螢幕的對角線長度為 6.12 吋。實際可視區較小。\r\n3.電池使用時間視使用情況及配置狀態而異。請參閱 apple.com tw batteries，以取得進一步資訊。\r\n4.須使用數據方案。
# 5g 僅於特定地區並透過特定電信業者提供。連線速度會因所在地情況及電信業者而異。如需 5g 支援的詳細資訊，請洽詢你的電信業者並參閱 apple.com tw iphone cellular。\r\n5.iphone 14 pro 具備防潑抗水與防塵功能，並且已 
# 在受控管的實驗室環境條件下測試，達到 iec 60529 標準的 ip68 等級 (在最深達 6 公尺水中最長可達 30 分鐘)。防潑抗水與防塵功能非永久狀態，並可能隨日常使用造成的耗損而下降。請勿為潮濕的 iphone 充電；請參閱使用手冊
# 了解清潔與乾燥處理說明。液體造成的損壞並不在保固範圍內。\r\n6.部分功能可能未在所有國家或地區提供。\r\n※相關注意事項，請見下方備註欄。",
#       "price": 36388,
#       "originPrice": 36388,
#       "author": "",
#       "brand": "",
#       "publishDate": "",
#       "isPChome": 1,
#       "isNC17": 0,
#       "couponActid": [
#         "C037767",
#         "C037739",
#         "C037740",
#         "C037768",
#         "C037769",
#         "C037742",
#         "C037770",
#         "C037743",
#         "C037771",
#         "C037745",
#         "C037746",
#         "C037772",
#         "C037747",
#         "C037773",
#         "C037774",
#         "C037749",
#         "C037775",
#         "C037748",
#         "C037776",
#         "C037750",
#         "C037751",
#         "C037777",
#         "C037752",
#         "C037778",
#         "C037753",
#         "C037779",
#         "C037754",
#         "C037780",
#         "C037781",
#         "C037755",
#         "C037782",
#         "C037756",
#         "C037783",
#         "C037757",
#         "C037758",
#         "C037784",
#         "C037759",
#         "C037785",
#         "C037786",
#         "C037766",
#         "C037787",
#         "C037760",
#         "C037788",
#         "C037761",
#         "C037762",
#         "C037789",
#         "C037790",
#         "C037763"
#       ],
#       "BU": "ec"
#     },
#     {
#       "Id": "DYAJ2Y-1900FW1K3",
#       "cateId": "DYAJ2Y",
#       "picS": "/items/DYAJ2Y1900FW1K3/000001_1691116286.jpg",
#       "picB": "/items/DYAJ2Y1900FW1K3/000001_1691116286.jpg",
#       "name": "Apple iPhone 14 Pro Max (256G)-深紫色(MQ9X3TA/A)",
#       "describe": "深紫色★送MFI傳輸線Apple iPhone 14 Pro Max (256G)-深紫色(MQ9X3TA/A)\\r\\n市價$42400．特價↘$３９６８８\r\n現買現省$2712\r\n限量特殺∼隨時回價∼！！\r\n------------------------------------------------------------------------\r\npchome安心保\r\n• pchome安心保 powered by applecare services\r\n需另外選購，購買 apple 商品後即可立即投保，\r\n請點此查詢價格及保障內容。\r\n• 6.7 吋超 retina xdr 顯示器， 
# 具備永遠顯示與 promotion 自動適應更新頻率技術\r\n• 動態島功能，靈活多變的 iphone 新玩法\r\n• 4800 萬像素主相機，解析度最高可達 4 倍\r\n• 電影級模式，現在支援最高可達 30 fps 的 4k 杜比視界\r\n• 動作模式，拍攝
# 手持影片又順又穩\r\n• 重大安全技術—車禍偵測，關鍵時刻為你撥打求救電話\r\n• 滿足一天所需的電池續航力，影片播放時間最長可達 29 小時\r\n• a16 仿生，極致的智慧型手機晶片。超飆速 5g 行動網路\r\n• 採用超瓷晶盾並且
# 防潑抗水，耐用性皆為業界領先\r\n• ios 16 帶來更多個性化設定、溝通與共享的新方式\r\n• ncc許可字號：ccai225g0100t2\r\n\r\n↑詳情請點圖片了解\r\n\r\n↑詳情請點圖片了解\r\n\r\n舊機不擔心，回收來pc\r\n↑舊愛不去新歡
# 不來\r\n\r\n法律聲明\r\n1.「sos 緊急服務」功能須透過行動網路連線或 wi-fi 通話使用。\r\n2.顯示器採用圓角設計。以矩形量測時，螢幕的對角線長度為 6.12 吋。實際可視區較小。\r\n3.電池使用時間視使用情況及配置狀態而
# 異。請參閱 apple.com tw batteries，以取得進一步資訊。\r\n4.須使用數據方案。5g 僅於特定地區並透過特定電信業者提供。連線速度會因所在地情況及電信業者而異。如需 5g 支援的詳細資訊，請洽詢你的電信業者並參閱 apple.com tw iphone cellular。\r\n5.iphone 14 pro 具備防潑抗水與防塵功能，並且已在受控管的實驗室環境條件下測試，達到 iec 60529 標準的 ip68 等級 (在最深達 6 公尺水中最長可達 30 分鐘)。防潑抗水與防塵功能非永久狀態，
# 並可能隨日常使用造成的耗損而下降。請勿為潮濕的 iphone 充電；請參閱使用手冊了解清潔與乾燥處理說明。液體造成的損壞並不在保固範圍內。\r\n6.部分功能可能未在所有國家或地區提供。\r\n※相關注意事項，請見下方備註欄。
# ",
#       "price": 39688,
#       "originPrice": 39688,
#       "author": "",
#       "brand": "",
#       "publishDate": "",
#       "isPChome": 1,
#       "isNC17": 0,
#       "couponActid": [],
#       "BU": "ec"
#     }

#endregion

#region 從PChome抓取URL回傳內容，取得區域:滑鼠右鍵->檢查->網路->篩選:all->名稱:results?

# 使用 askstring函數 彈出filedialog對話框前來輸入搜尋內容
target = simpledialog.askstring("輸入搜尋內容", "請輸入搜尋內容:")

if target is not None: 
    # 創建UserAgent物件
    ua = UserAgent()
    # 從UserAgent物件中獲取一個隨機的User-Agent，用於發送HTTP請求，降低被網站阻擋或限制的風險
    user_agent = ua.random
    my_head = {'user-agent': user_agent}
    print(f'\r\nmy_head : {my_head}')

    get_values = [] #取得每頁搜尋到的資料總攬

    # 取得前三頁資料
    for page in range(1,4): 
        
        page_url = f'https://ecshweb.pchome.com.tw/search/v3.3/all/results?q={target}&page={page}&sort=rnk/dc'
        page_response = requests.get(page_url, headers= my_head)

        page_content = page_response.content.decode(encoding='utf-8')
        # 將JSON轉換成Python字典
        page_res_data = json.loads(page_content)
        # 使用json.dumps()將字典轉換回JSON格式，indent參數指定縮排，ensure_ascii=False則允許輸出非ASCII字符，這樣中文字符就可以正確顯示
        pretty_json = json.dumps(page_res_data, indent=2, ensure_ascii=False) 
        #print(pretty_json)

        # 設定每頁只顯示20筆
        results_per_page = 20

        if page == 1:

            # 搜尋總筆數
            total_rows = page_res_data['totalRows']
            print(f'搜尋總筆數:{total_rows}')

            # 抓取總頁數
            pages_count = int(total_rows / results_per_page)

            if int(total_rows % results_per_page) != 0:
                pages_count += 1
            print(f'所需頁數:{pages_count}\r\n')

        # 回傳url
        print(page_response.url + '\r\n')

        icount = len(page_res_data['prods'])
        print(f'第 {page} 頁 ; 共{icount}筆')    
        # 提取商品名稱和價格
        for product in page_res_data['prods']:
            name = product['name']
            price = product['price']
            get_values.append(f'商品名稱: {name}, 價格: {price}')
            print(f'商品名稱: {name}, 價格: {price}')
            # step3 把這些籃子轉成dataframe
        print("\n")

    print('--------------------------------------------------------------------------------------------------\r\n')

    if get_values is not None:  

        # 將 get_values 資料轉換成 DataFrame
        data = [item.split(', ') for item in get_values]
        df = pd.DataFrame(data, columns=['商品名稱', '價格']) 

        # 創建一個隱藏的 root 視窗以使用檔案對話框
        root = Tk()
        root.withdraw()
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", ('*.xls', '*.xlsx'))], title='另存新檔')
        
        if file_path:  
            # 判斷使用者選擇的副檔名，並將 DataFrame 儲存至相應的 Excel 檔案
            if file_path.split('.')[1] == 'xlsx':
                # 使用 pd.ExcelWriter 創建一個寫入器對象，並在其中儲存 DataFrame
                # 在with區塊中，你可以使用writer物件來寫入資料到Excel檔案中的各個工作表。一旦with區塊結束，pd.ExcelWriter會自動關閉並保存Excel檔案，確保寫入操作的正確性和完整性
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='商品價目')
                    workbook = writer.book
                    worksheet = writer.sheets['商品價目']
                    
                    # 設定商品名稱與價格欄位的寬度
                    column_widths = [max(df[col].astype(str).apply(len).max(), len(col)) for col in df.columns]
                    for i, width in enumerate(column_widths):
                        worksheet.column_dimensions[get_column_letter(i + 1)].width = width

            else:  
                df.to_excel(file_path, index=False)
                
        print(f'資料已儲存至文字檔：{file_path}')

else:
        print("使用者取消操作")

#endregion


