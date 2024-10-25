import logging
import pandas as pd
from datetime import datetime, timedelta
import time

from selenium import webdriver
from selenium.webdriver.support.select import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common import exceptions
import pandas as pd
import glob

from setting_wiz import *
from map_rac import *
import setting_wiz


def login():
    try:
        option = webdriver.ChromeOptions()
        pref = {'download.default_directory': Data_path}
        option.add_experimental_option('prefs', pref)
        option.add_argument('ignore-certificate-errors')
        option.add_argument("--no-sandbox")
        option.add_experimental_option("detach", True)
        option.add_argument("--disable-dev-shm-usage")
        driver = webdriver.Chrome(options= option)
        driver.implicitly_wait(30)
    except Exception as e:
        print(e)
    
    login_ins = setting_wiz.Login()
    login_ins.login()
    login_ins.webdriver()
    driver.get(login_ins.wiz)
    print("Open webdriver")
    driver. maximize_window()
    driver.find_element(By.NAME, 'account').send_keys(login_ins.user)
    driver.find_element(By.NAME, 'password').send_keys(login_ins.password)
    
    element = driver.find_element(By.CSS_SELECTOR, 'div.handler.iconfont-menu.icon-qianjin')
    offset_x = 350
    actions = ActionChains(driver)
    time.sleep(2)
    actions.click_and_hold(element).perform()
    time.sleep(2)
    actions.move_by_offset(offset_x, 0).perform()
    time.sleep(2)
    actions.release().perform()
    
    driver.find_element(By.XPATH, '/html/body/div[3]/div/div[2]/div[2]/div/div[2]/div/div/div/span[2]').click()
    driver.find_element(By.XPATH, '/html/body/div[3]/div/div[3]/span/button[2]/span').click()
    return driver

def search_master(driver, date):
    # current_date = datetime.now() - timedelta(days=1)
    start_date = date.strftime("%d/%m/%Y")
    # start_date = "30/09/2024"
    driver.implicitly_wait(300)
    ai_outbound_element = driver.find_element(By.XPATH, "//div[@class='el-tooltip navBar-navItem-overflow']/span[text()=' AI Outbound']")
    ai_outbound_element.click()
    time.sleep(2)
    outbound_call_el = driver.find_element(By.XPATH, "//div[@class='el-tooltip navBar-navItem-overflow']/span[text() =' Call Records']")
    outbound_call_el.click()
    time.sleep(2)
    filter_button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'filter__btn')]"))
    )
    driver.execute_script("arguments[0].click();", filter_button)
    time.sleep(4)
    call_time_input = driver.find_element(By.XPATH, "//input[@placeholder='Select time']")
    call_time_input.click()
    time.sleep(3)
    call_date_input = driver.find_element(By.XPATH, "//input[@placeholder='Select date']")
    call_date_input.click()
    time.sleep(2)
    call_date_input.clear()
    call_date_input.send_keys(start_date + Keys.ENTER)
    # driver.execute_script("arguments[0].value = arguments[1];", call_date_input, start_date)
    # call_date_input.send_keys(Keys.ENTER)
    js = """
            $(document).ready(function(){
                //หาสแปน ok แล้วหาปุ่มที่ใกบ้สุดของ ok กดแม่งโลดด
                $("span:contains('OK')").closest("button").click();
            });
        """
    driver.execute_script(js)
    time.sleep(2)
    call_time_input_second = driver.find_elements(By.XPATH, "//input[@placeholder='Select time']")[1]
    call_time_input_second.click()
    call_date_input_second = driver.find_elements(By.XPATH, "//input[@placeholder='Select date']")[1]
    call_date_input_second.click()
    time.sleep(2)
    call_date_input_second.clear()
    call_date_input_second.send_keys(start_date + Keys.ENTER)
    driver.execute_script(js)
    js_filter = """
            $(document).ready(function(){
                $("span:contains('Apply Filters')").closest("button").click(); 
            });
        """    
    driver.execute_script(js_filter)
    js_export = """
            $(document).ready(function() {
                // ค้นหาและคลิกปุ่มแรกที่เปิดป๊อปอัพ
                $("span:contains('Export')").closest("button").click();

                // รอให้ป๊อปอัพโหลดเสร็จ
                setTimeout(function() {
                    $("button:contains('Export')").filter(".el-button--primary").click();
                }, 2000);
            }); 
         """
    driver.execute_script(js_export)
        
def Download_Data(driver):
    max_attempts = 10
    attempt = 0
    while attempt < max_attempts:
        try:
            time.sleep(10)
            # ค้นหาอิลิเมนต์ที่ต้องการ
            # download_button = driver.find_element(By.XPATH, "/html/body/div[9]/div/div[2]/div/div[1]/div[4]/div[2]/table/tbody/tr[1]/td[7]/div/span/button[1]")
            download_button = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[9]/div/div[2]/div/div[1]/div[4]/div[2]/table/tbody/tr[1]/td[7]/div/span/button[1]"))
            )
            # ตรวจสอบข้อความภายในอิลิเมนต์
            button_text = download_button.text.strip()
            if "Download" in button_text:  # ถ้าข้อความเป็น "Download"
                driver.execute_script("arguments[0].click();", download_button)
                print(f"Click successful on attempt {attempt + 1}")
                break
            else:
                print(f"Button text is '{button_text}' on attempt {attempt + 1}, waiting for 'Download'")
                attempt += 1
        except Exception as e:
            print(e)
            print(f"Error on attempt {attempt + 1}: {e}")
            attempt += 1

    if attempt == max_attempts:
        print("Max attempts reached. The button text never changed to 'Download'.")

    # ตรวจสอบว่า jQuery ยังคงมี request active อยู่หรือไม่
    wait = 1
    while wait == 1:
        wait = driver.execute_script('return jQuery.active;')  # ถ้าไม่มี active requests, jQuery.active จะเป็น 0
        time.sleep(0.5)

    time.sleep(5)
    logging.debug('Downloading')

    # ตั้งเวลารอเพิ่มเติมเพื่อให้การดาวน์โหลดเสร็จสมบูรณ์
    download_complete = False
    for _ in range(30):  # ลองเช็คการดาวน์โหลดทุกๆ 2 วินาที นานสุด 1 นาที
        check_xlsx = glob.glob(os.path.join(Data_path, '*.xlsx'))
        if check_xlsx:
            download_complete = True
            break
        time.sleep(2)

    if download_complete:
        print('Download Success.')
    else:
        print('Download failed')

    logging.info('Quit web driver')
    driver.close()

    
def get_data(date):
    dr = login()
    time.sleep(5)
    search_master(dr, date)
    time.sleep(4)
    Download_Data(dr)
    
if __name__ == '__main__':
    current_date = datetime.now() - timedelta(days=1)
    dr = login()
    time.sleep(3)
    search_master(dr, current_date)
    time.sleep(4)
    Download_Data(dr)