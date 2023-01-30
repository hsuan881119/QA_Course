from appium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os

def setup_phone_and_driver(url):
    desired_caps = {
        'platformName': 'Android',
        'platformVersion': '9',
        'deviceName': 'ASUS_X01BDA',
        'autoGrantPermissions': True,
        'browserName': 'Chrome',
        'noReset': False,
        'automationName': 'UiAutomator2'
    }
    driver = webdriver.Remote('http://localhost:4723/wd/hub', desired_caps)
    driver.get(url)
    return driver

def wait(driver, path):
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, path)))

def screenshot(driver, filename):
    time.sleep(0.2)
    driver.get_screenshot_as_file(os.getcwd() + filename)

def wait_and_click(driver, path):
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, path)))
    goal = driver.find_element(By.XPATH, path)
    goal.click()
    return goal

def mv_to_stopissuing_cards(driver, path):
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight-1500);")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, path)))
    stopissuing_cards_btn = driver.find_element(By.XPATH, path)
    stopissuing_cards_btn.click()

def find_items(driver, list_path, items_path):
    wait(driver, list_path)
    list = driver.find_element(By.XPATH, list_path)
    items = list.find_elements(By.XPATH, items_path)
    return items


#get driver
driver = setup_phone_and_driver(url='https://www.cathaybk.com.tw/cathaybk/')
#loading complete page and screenshot
wait(driver, './/div[@class="cubre-o-quickLink"]')
screenshot(driver, '/home_page.png')
#click menu>product_introduction>credit_card
menu_btn = wait_and_click(driver, './/div[@class="cubre-o-header__burger"]')
product_intro_btn = wait_and_click(driver, '/html/body/div[1]/header/div/div[3]/div/div[2]/div/div/div[1]/div[1]')
creditcard_btn = wait_and_click(driver, './/div[@class="cubre-o-menuLinkList__btn"]')
#find list items and screenshot
creditcardlist_items = find_items(driver, 
    '/html/body/div[1]/header/div/div[3]/div/div[2]/div/div/div[1]/div[2]/div/div[1]/div[2]', 
    './/a[@class="cubre-a-menuLink"][@id="lnk_Link"]')
screenshot(driver, '/creditcard_list.png')
#click card_intro
card_intro_btn = creditcardlist_items[0]
card_intro_btn.click()
time.sleep(10)
#mv_to_stopissuing_cards(driver, '/html/body/div[1]/main/article/div/div/div/div[1]/div/div/a[6]')
#find stopissuing_cards items and screenshot
stopissuing_cards_items = find_items(driver,
    '/html/body/div[1]/main/article/section[6]/div/div[2]/div/div[2]',
    './/span[contains(@class, "swiper-pagination-bullet")]')   
for (idx, item) in enumerate(stopissuing_cards_items):
    item.click()
    wait(driver, ('/html/body/div[1]/main/article/section[6]/div/div[2]/div/div[1]/div['+str(idx+1)+']'))
    screenshot(driver, ('/stopissuing_card_'+str(idx+1)+'.png'))
#output results
print("-----------------------------------")
print("信用卡選單項目數量:",len(creditcardlist_items))
num_of_stopissuing_cards = 0
for path in os.listdir(os.getcwd()):
    if os.path.isfile(os.path.join(os.getcwd(), path)) and path.startswith("stopissuing_card_"):
        num_of_stopissuing_cards += 1
print("-----------------------------------")
if num_of_stopissuing_cards == len(stopissuing_cards_items):
    print("停發卡數量與截圖數量一致,皆為:",len(stopissuing_cards_items))
print("-----------------------------------")