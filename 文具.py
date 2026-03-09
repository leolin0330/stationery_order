# 匯入套件
import pandas as pd          # 用來讀取 Excel
import time                  # 控制程式等待時間
from selenium import webdriver   # Selenium 瀏覽器自動化
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By   # 用來定位網頁元素
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait   # 等待元素出現
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoAlertPresentException, UnexpectedAlertPresentException
from webdriver_manager.chrome import ChromeDriverManager


# ================================
# 初始化瀏覽器
# ================================
def init_driver():

    chrome_options = Options()

    # detach=True 代表程式結束後瀏覽器不會自動關閉（方便除錯）
    chrome_options.add_experimental_option("detach", True)

    # 自動下載對應版本的 ChromeDriver
    service = Service(ChromeDriverManager().install())

    # 啟動 Chrome 瀏覽器
    return webdriver.Chrome(service=service, options=chrome_options)


# ================================
# 登入系統
# ================================
def login(num_reg, account, password, driver):

    # 進入登入頁面
    url = 'http://advantage.stapro.com.tw/Powerlink/form/login.aspx'
    driver.get(url)

    # 輸入公司統編
    driver.find_element(By.ID, "txtKHBH").send_keys(num_reg)

    # 輸入帳號
    driver.find_element(By.ID, "txtUID").send_keys(account)

    # 輸入密碼
    driver.find_element(By.ID, 'txtPassword').send_keys(password)

    # 偵測是否有 alert 彈出視窗
    try:
        WebDriverWait(driver, 5).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        print("彈出訊息:", alert.text)
        alert.accept()

    except:
        print("❌ 沒有彈出視窗，繼續執行程式")

    # 點擊登入按鈕
    driver.find_element(By.ID, 'ibtLogin').click()


# ================================
# 進入訂購頁面
# ================================
def enter_order_page(driver):

    # 等待登入成功後按鈕出現
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'But_Login'))
    ).click()

    # 處理可能出現的 alert
    try:
        WebDriverWait(driver, 5).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        print("彈出訊息:", alert.text)
        alert.accept()

    except:
        print("❌ 沒有彈出視窗")


# ================================
# 搜尋商品
# ================================
def search_product(driver, product_code):

    try:

        # 找到搜尋框
        search_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'ctl00_WucSearch1_txtTerm'))
        )

        # 清空搜尋框
        search_box.clear()

        # 輸入產品編號
        search_box.send_keys(product_code)

        # 點擊搜尋
        driver.find_element(By.ID, 'ctl00_WucSearch1_btnSearch').click()

        # 等待商品頁面出現數量欄位
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'ctl00_ContentPlaceHolder1_dlsOrderBY_ctl00_tbxAmount'))
        )

        print(f"✅ 產品 {product_code} 找到")
        time.sleep(1)

        return True

    # 如果商品不存在會出現 alert
    except UnexpectedAlertPresentException:

        alert = driver.switch_to.alert
        print(f"⚠️ 找不到產品 {product_code}：{alert.text}")
        alert.accept()

        return False

    except Exception as e:

        print(f"❌ 搜尋產品 {product_code} 時發生錯誤：{e}")
        return False


# ================================
# 加入購物車
# ================================
def add_to_cart(driver, quantity):

    try:

        # 找到數量欄位
        quantity_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'ctl00_ContentPlaceHolder1_dlsOrderBY_ctl00_tbxAmount'))
        )

        quantity_input.clear()

        # 輸入購買數量
        quantity_input.send_keys(quantity)

        # 點擊加入購物車
        driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_dlsOrderBY_ctl00_ibtInsertCart').click()

        # 等待成功訊息
        WebDriverWait(driver, 5).until(EC.alert_is_present())

        alert = driver.switch_to.alert
        print("🛒 加入購物車成功：" + alert.text)
        alert.accept()

    except:

        print("⚠️ 加入購物車可能失敗")


# ================================
# 從 Excel 讀取訂單
# ================================
def process_order_from_excel(driver, file_path, failed_orders_file):

    # 讀取 Excel
    data = pd.read_excel(file_path)

    failures = []

    # 逐筆處理商品
    for _, order in data.iterrows():

        product_code = str(order['產品編號'])
        quantity = str(order['數量'])

        print(f"\n🔍 處理商品：{product_code}")

        try:

            # 搜尋商品
            if search_product(driver, product_code):

                # 加入購物車
                add_to_cart(driver, quantity)

            else:

                # 商品不存在記錄
                failures.append({
                    '產品編號': product_code,
                    '數量': quantity,
                    '原因': '未找到商品'
                })

        except Exception as e:

            failures.append({
                '產品編號': product_code,
                '數量': quantity,
                '原因': '處理錯誤'
            })

    # 如果有失敗訂單輸出 Excel
    if failures:

        pd.DataFrame(failures).to_excel(failed_orders_file, index=False)

        print(f"\n❗ 有失敗項目，已匯出至 {failed_orders_file}")

    else:

        print("\n✅ 所有商品成功加入購物車！")


# ================================
# 前往購物車
# ================================
def send_order(driver):

    driver.find_element(By.ID, 'ctl00_WucLogo1_hlShopingCart').click()

    try:

        WebDriverWait(driver, 5).until(EC.alert_is_present())
        alert = driver.switch_to.alert

        print("🧾 彈出訊息:", alert.text)
        alert.accept()

    except:

        print("❌ 沒有彈出視窗")


# ================================
# 主程式
# ================================
def main():

    # 啟動瀏覽器
    driver = init_driver()

    # 登入系統
    login('42612538', 'D000', '1111', driver)

    # 進入訂購頁
    enter_order_page(driver)

    # Excel 檔案
    file_path = 'product.xlsx'

    # 失敗訂單紀錄
    failed_orders_file = '錯誤清單.xlsx'

    # 執行自動下單
    process_order_from_excel(driver, file_path, failed_orders_file)

    # 前往購物車
    send_order(driver)

    input("\n📦 按 Enter 關閉瀏覽器")

    driver.quit()


# 執行程式
if __name__ == "__main__":
    main()




    # driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_btnOrder').click()
    #
    # # 彈出視窗按確認
    # try:
    #     WebDriverWait(driver, 5).until(EC.alert_is_present())  # 最多等 5 秒
    #     alert = driver.switch_to.alert
    #     print("彈出訊息:", alert.text)  # (可選) 打印彈出訊息
    #     alert.accept()  # 點擊【確定】

    # except:
    #     print("❌ 沒有彈出視窗，繼續執行程式")