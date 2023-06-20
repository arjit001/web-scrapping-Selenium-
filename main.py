from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import time
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import textwrap
opts = webdriver.ChromeOptions()
opts.add_argument("--start-maximized")
# opts.add_argument("--window-size=1115,649")
# to disable bot detector
# opts.add_experimental_option("excludeSwitches", ["enable-automation"])
# opts.add_experimental_option('useAutomationExtension', False)
# For ChromeDriver version 79.0.3945.16 or over
# opts.add_argument('--disable-blink-features=AutomationControlled')
driver = webdriver.Chrome(ChromeDriverManager().install(), options=opts)

driver.get('https://shopla.homelegance.com/nmember/tablePageInit.htm?websiteId=1')
time.sleep(1)

# a = input('press enter')

driver.find_element(By.XPATH, '//*[@id="username"]').send_keys('****@***')
driver.find_element(By.XPATH, '//*[@id="password"]').send_keys('******')
driver.find_element(By.XPATH, '/html/body/div/div[5]/div/div/div/div[2]/div/form/div/div[2]/input[2]').click()
time.sleep(1)
# driver.find_element(By.XPATH, '/html/body/div[1]/div[9]/div/div/div[2]/div[2]/div/div[2]/div/div/div[3]/div/a').click()

# a = input('press enter')

file = pd.read_excel('AshleyDirect.xlsx')

filename = 'fetched.csv'



with open(filename, mode='w+', encoding='utf-8-sig') as f:
    f.write(
        '{}, {}, {}, {} , {}, {}, {}, {}, {}, {}, {}\n'.format('Item Number', 'ProName', 'ProDisc', 'OverallWeight', 'General Dimensions', 'Carton Dimension', 'CartWeight', 'Image',
                                          'Assembly Instruction', 'Bullet', 'Price'))

for i in range(0, len(file)):
    sku = file['Item Number'][i]
    # del ProName, ProDisc, OverallWeight, GenDimension, cartonDim, Cartweight, image, assemblyIns, bullet, price
    driver.find_element(By.XPATH, '//*[@id="web_search_model"]').clear()
    driver.find_element(By.XPATH, '//*[@id="web_search_model"]').send_keys(sku)
    # time.sleep(1)
    driver.find_element(By.XPATH, '//*[@id="web_search_model"]').send_keys(Keys.RETURN)
    time.sleep(2)
    # webdriver.Keys.RETURN
    try:
        driver.find_element(By.XPATH, '//*[@id="categoryModelList_vue"]/div[2]/div/div[1]/div[1]').click()
        # driver.find_element(By.XPATH, '//*[@id="categoryModelList_vue"]/div[2]/div/div[1]/div[1]/div[2]/a/div/img').click()

        time.sleep(1)



        # a = input('press enter')
        # driver.find_element(By.XPATH, '//*[@id="categoryModelList_vue"]/div[2]/div/div[1]/div[1]/div[2]/a/div/img').click()
        # time.sleep(5)



        try:
            ProDisc = driver.find_element(By.XPATH, '//*[@id="pic_1"]').get_attribute('src')
            # time.sleep(1)
        except:
            ProDisc = 'not found'

        try:
            OverallWeight = driver.find_element(By.XPATH, '//*[@id="pic_2"]').get_attribute('src')
            # time.sleep(1)
        except:
            OverallWeight = 'not found'

        try:
            GenDimension = driver.find_element(By.XPATH, '//*[@id="pic_3"]').get_attribute('src')
            # time.sleep(1)
        except:
            GenDimension = 'not found'
        try:
            Cartweight = driver.find_element(By.XPATH, '//*[@id="pic_4"]').get_attribute('src')
            # time.sleep(1)
        except:
            Cartweight = 'not found'
        try:
            image = driver.find_element(By.XPATH, '//*[@id="pic_5"]').get_attribute('src')
            # time.sleep(1)
        except:
            image = 'not found'
        try:
            cartonDim = driver.find_element(By.XPATH, '//*[@id="pic_6"]').get_attribute('src')
        except:
            cartonDim = 'not found'
        try:
            assemblyIns = driver.find_element(By.XPATH, '//*[@id="pic_7"]').get_attribute('src')
        except:
            assemblyIns = 'not found'
        try:
            bullet = driver.find_element(By.XPATH, '//*[@id="product_details_content"]/ul').text.replace("\n", "@")
        except:
            bullet = 'not found'
        try:
            price = driver.find_element(By.XPATH, '//*[@id="modal_product_price"]').text
        except:
            price = 'not found'

        try:
            # driver.find_element(By.XPATH, '//*[@id="model_detail_vue"]/div[1]/div/div/div[3]/div[6]/div[3]/p').click()
            ProName = driver.find_element(By.XPATH, '//*[@id="pic_0"]').get_attribute('src')
            # time.sleep(1)
        except:
            ProName = 'not found'

            # time.sleep(1)
        # driver.find_element(By.XPATH, '//*[@id="myModal"]/div/div/div[1]/button').click()
    # driver.find_element(By.XPATH, '//*[@id="product_details_content"]/ul/li[7]').send_keys(Keys.ESCAPE)
    except:
        ProName, ProDisc, Ov    erallWeight, GenDimension, cartonDim, Cartweight, image, assemblyIns, bullet, price = 'not found', 'not found', 'not found', 'not found', 'not found', 'not found', 'not found', 'not found', 'not found', 'not found'

    with open(filename, 'a', encoding='utf-8-sig') as f:
        f.write('{}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}\n'.format(sku, ProName, ProDisc, OverallWeight, GenDimension, cartonDim, Cartweight, image, assemblyIns, bullet, price))
    # a = input('press enter')

