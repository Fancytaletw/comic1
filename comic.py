# 使用規則:
# 請將下面"可變動變數建立"填寫完畢
# 請將漫畫總集數填入total_episode_amount
# 請將漫畫儲存的資料夾指定，最後PDF檔也會在那裡

from selenium import webdriver 
from time import sleep, time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import wget
from PIL import Image
import img2pdf
import shutil

# 清理資訊欄
os.system("cls")

# 可更動變數建立
comic_web_URL = "https://mh5.tw/series-wunydlt-412076-1-%E4%BE%86%E8%87%AA%E6%B7%B1%E6%B7%B5" # 漫畫第一頁網址
total_episode_amount = 52 # 總集數
save_folder_location = "D:/" # 確定資料夾儲存位置
img_save_folder_name = "來自深淵" #確定資料夾名稱
if_accept_term = True # 是否需處理同意條款
if_save_img = False # 下載的圖片要不要儲存
comic_name = "來自深淵"

# 不可更動變數建立
WebDriver_path = "D:/edgedriver_win64/msedgedriver.exe"
WebDriver=webdriver.Edge(WebDriver_path)
total_pages_per_episode = 0
set_size = 1500
temp_save_folder_location = save_folder_location + "temp_save/"
save_folder_location = save_folder_location+img_save_folder_name+"/"
if_exception_handling_continue = True
error_times_counter = 0
start_time = time()

# 建立目標資料夾
if os.path.isdir(save_folder_location):
    shutil.rmtree(save_folder_location)
os.mkdir(save_folder_location)

# 若漫畫要儲存，建立另一個暫時性資料夾
if if_save_img:
    if os.path.isdir(temp_save_folder_location[:len(temp_save_folder_location)-1]):
        shutil.rmtree(temp_save_folder_location[:len(temp_save_folder_location)-1])
    os.mkdir(temp_save_folder_location[:len(temp_save_folder_location)-1])

# 下載漫畫
WebDriver.get(comic_web_URL)
WebDriver.maximize_window()

# 處理同意條款
if if_accept_term:
    element = WebDriverWait(WebDriver, 5).until(
            EC.presence_of_element_located((By.CLASS_NAME, "light"))
        )
    element.click()

# 開始下載漫畫
for episode in range(1,total_episode_amount+1):
    # 一集漫畫所有的圖片位置
    sorted_comic_img_path_list = []
    # 抓取一本書有幾頁
    total_pages_per_episode = WebDriver.find_element(By.XPATH, "/html/body/div/header/div/div/div[3]/h2")
    total_pages_per_episode = str(total_pages_per_episode.text)
    total_pages_per_episode = int(total_pages_per_episode[total_pages_per_episode.index("(")+1:-2])
    all_img_path_list = []

    # 計算每一集裡面的頁數
    comic_page_processing = 1
    while (comic_page_processing<=total_pages_per_episode):
        sleep(2)
        # 抓取圖片URL
        if_exception_handling_continue = True
        error_times_counter = 0
        while (if_exception_handling_continue):
            try:
                img_link = WebDriver.find_element(By.XPATH,"/html/body/div/div[4]/div/img")
                img_link = img_link.get_attribute("src")
                if_exception_handling_continue = False
            except:
                error_times_counter+=1
                if error_times_counter>=5:
                    print("正在連結到第"+str(episode)+"集第"+str(comic_page_processing)+"頁時遇到困難，將重整並於4秒後重試")
                    WebDriver.refresh()
                    error_times_counter = 0
                else:
                    print("正在連結到第"+str(episode)+"集第"+str(comic_page_processing)+"頁時遇到困難，將於4秒後重新開始")
                sleep(4)
        
        # 命名並下載圖片
        save_at = os.path.join(save_folder_location,str(episode)+"-"+str(comic_page_processing)+".jpg")
        if_exception_handling_continue = True
        error_times_counter = 0
        while (if_exception_handling_continue):
            try:
                print("\n"+str(episode)+"-"+str(comic_page_processing)+".jpg"+" 正在下載")
                print("=============================================")
                wget.download(img_link,save_at)
                if_exception_handling_continue = False
            except Exception:
                error_times_counter+=1
                if (error_times_counter>=4):
                    print("下載第"+str(episode)+"集第"+str(comic_page_processing)+"頁時發生錯誤，將重整並於3秒後重試")
                    WebDriver.refresh()
                    error_times_counter = 0
                    sleep(3)
                else:
                    print("下載第"+str(episode)+"集第"+str(comic_page_processing)+"頁時發生錯誤，將於2秒後重試")
                    sleep(2)
            
        # 處理圖片
        print("\n============開始處理"+str(episode)+"-"+str(comic_page_processing)+".jpg"+"============")
        
        img_obj = Image.open(save_at)
        # 翻轉圖片
        img_size = img_obj.size
        if img_size[0]>img_size[1]:
            img_obj = img_obj.rotate(270,expand=1)
            img_size = (img_size[1],img_size[0])
        # 處理成等高
        img_enlarge_rate = set_size/img_size[1]
        img_obj = img_obj.resize((int(img_size[0]*img_enlarge_rate),set_size))

        # 圖片是否要儲存
        if if_save_img:
            save_at = os.path.join(temp_save_folder_location,str(episode)+"-"+str(comic_page_processing)+".jpg")
        else:
            os.remove(save_at)
        sorted_comic_img_path_list.append(save_at)
        img_obj.save(save_at)
        img_obj.close()
        
        print(str(episode)+"-"+str(comic_page_processing)+".jpg"+"處理完畢")

        # 點擊下一頁
        if_exception_handling_continue = True
        error_times_counter = 0
        while (if_exception_handling_continue):
            try:
                next_page = WebDriver.find_element(By.LINK_TEXT, str(comic_page_processing+1))
                next_page.click()
                if_exception_handling_continue = False
            except Exception:
                error_times_counter+=1
                if error_times_counter>=3:
                    print("點擊第"+str(episode)+"集第"+str(comic_page_processing+1)+"頁時遇到困難，將重整並於4秒後並重試")
                    WebDriver.refresh()
                    error_times_counter = 0
                else:
                    print("點擊第"+str(episode)+"集第"+str(comic_page_processing+1)+"頁時遇到困難，將於4秒後重試")
                sleep(4)
        comic_page_processing+=1

    # 開始製作PDF
    print("============第"+str(episode)+"集所有圖片下載完成============")
    print("============開始製作"+comic_name+"第"+str(episode)+"集PDF============\n")
    gonna_convert_PDF_path = os.path.join(save_folder_location, "來自深淵第"+str(episode)+"集.pdf")
    if os.path.isfile(gonna_convert_PDF_path):
        os.remove(gonna_convert_PDF_path)
    with open(gonna_convert_PDF_path, 'wb') as file:
        file.write(img2pdf.convert(sorted_comic_img_path_list))
        file.close()
        
    print("============"+comic_name+"第"+str(episode)+"集PDF製作完畢============\n")
    sleep(6)
    # 清除暫時資料夾內的資料
    for img_path in sorted_comic_img_path_list:
        os.remove(img_path)

    # 點擊下一集
    if_exception_handling_continue = True
    error_times_counter = 0
    while (if_exception_handling_continue):
        try:
            next_episode = WebDriver.find_element(By.XPATH,"/html/body/div/div[7]/div/div[1]/div[4]/a")
            next_episode.click()
            if_exception_handling_continue = False
        except Exception:
            error_times_counter+=1
            if error_times_counter>=3:
                print("點擊進入第"+str(episode+1)+"集時發生錯誤，將重整並於4秒後重試")
                WebDriver.refresh()
                error_times_counter = 0
            else:
                print("點擊進入第"+str(episode+1)+"集時發生錯誤，將於4秒後重試")
            sleep(4)   

WebDriver.close() # 關閉網站視窗

# 若漫畫要儲存，刪除暫時性資料夾
if if_save_img and os.path.isdir(temp_save_folder_location[:len(temp_save_folder_location)-1]):
    shutil.rmtree(temp_save_folder_location[:len(temp_save_folder_location)-1])

# 計算經過時間
end_time = time()
print("花費時間: " + str(int((end_time-start_time)//3600)) + "hr "+str(int(((end_time-start_time)-(end_time-start_time)//3600*3600)//60))+\
    "min " + str(int((end_time-start_time) - (end_time-start_time)//60*60)) + "sec")


print('============所有程式已結束============')