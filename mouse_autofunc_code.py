import pyautogui as pag
from pynput import keyboard
import msvcrt
import time
import pyperclip

#----------------------------判定鍵盤按下的按鍵----------------------------
def on_press(key):
    try:
        print('按下按鍵：{0}'.format(key.char))
    except AttributeError:
        print('按下特殊按鍵：{0}'.format(key))

def on_release(key):
    if key == keyboard.Key.enter:
        print('使用者按下了 enter 鍵，程式結束。')
        return True
    else:
        return False
#-------------------------------------------------------------------------
    

screen_w,screen_h = pag.size()
print(f'螢幕尺寸 => 寬:{screen_w} 高:{screen_h}')

#參考文獻:https://yanwei-liu.medium.com/pyautogui-%E4%BD%BF%E7%94%A8python%E6%93%8D%E6%8E%A7%E9%9B%BB%E8%85%A6-662cc3b18b80

#------------------------------取得鼠標座標位置--------------------------------
# while True:
#     print('鼠標位置判定:',pag.position())

#     if msvcrt.kbhit(): #用於检查是否有键盘按键输入
#         key = msvcrt.getch() #獲取按鍵的值ASCII
#         if ord(key) == 13:  # 13 is the ASCII code for "Enter" key
#             print('最終座標位置:',pag.position())
#             break
#------------------------------------------------------------------------------

#----------------------------自動控制另存圖片流程操作----------------------------
#typewrite函数用於模擬键盘输入，但它無法處理非ASCII字符，例如"狗"这个中文字符，所以改用'pyperclip.copy+pag.hotkey'此方法來實現輸入'狗'
#執行過程中請勿亂控制滑鼠與鍵盤，會打亂程式操作流程

#duration :用num_seconds秒的时间把鼠標移动到(x, y)位置
num_seconds = 2

#點選瀏覽器:x=419 y=1056
pag.moveTo(416, 1056, duration=num_seconds)
pag.leftClick()

#開新分頁:x=300 y=12
pag.leftClick(300, 12, duration=num_seconds)

#滑鼠左鍵按google選項:x=465 y=253
pag.leftClick(x=465, y=253, duration=num_seconds)

#滑鼠左鍵按搜索位置:x=693 y=505
pag.leftClick(x=693, y=505, duration=num_seconds)

# 將"狗"複製到搜索處
pyperclip.copy('狗')
# 模擬Ctrl+V操作来输入中文字符
pag.hotkey('ctrl', 'v')
time.sleep(2)
#按下ENTER键
pag.press('enter')

#點擊圖片分類 x=259 y=183
pag.leftClick(x=259, y=183, duration=num_seconds)

#移至要儲存的圖片 x=159 y=434 並滑鼠右鍵
pag.moveTo(159, 434, duration=num_seconds)
pag.rightClick()

#另存圖片 x=207 y=684
pag.leftClick(x=207, y=684, duration=num_seconds)

#移至檔名設定 x=170 y=430
pag.leftClick(170, 430, duration=num_seconds)

#輸入檔名
pag.hotkey('ctrl', 'a')
pyperclip.copy('圖狗1')
# 模擬Ctrl+V操作来输入中文字符
pag.hotkey('ctrl', 'v')
time.sleep(2)
#按下ENTER键
pag.press('enter')
time.sleep(2)

#確保有無出現'覆蓋圖片'對話框
pag.press('y')

#返回google主頁
pag.leftClick(x=64, y=127, duration=num_seconds)
#-------------------------------------------------------------------------
