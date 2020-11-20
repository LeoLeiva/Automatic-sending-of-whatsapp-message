from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import time
import xlrd 

# This is the name of the excel template
excel = ("outcome.xlsx") 

wb = xlrd.open_workbook(excel) 
sheet = wb.sheet_by_index(0) 

ChromeOptions = webdriver.ChromeOptions()
ChromeOptions.add_argument('--no-sandbox') 
# Here save the cokies of the chromedriver and sesion whatsapp
ChromeOptions.add_argument('--user-data-dir=D:\\01\\01data')
driver = webdriver.Chrome('chromedriver.exe', options=ChromeOptions)
driver.get("https://web.whatsapp.com")

# If you need setting qr on whatsapp web, change the time to 30 or more seconds
time.sleep(5)


for e in range(sheet.nrows):
    name = str(sheet.cell_value(e,0).title())

    # If you put the countr code on the excel, you can clear this line "countrycode = str("+54")"
    countrycode = str("+54")
    phone = int(sheet.cell_value(e,1))
    outcome = str(sheet.cell_value(e,2).title())

    # Here you can change de text that is send
    text = "Hola {} tu resultado del hisopado es {}".format(name,outcome)

    # If you want the country code you can go in the excel spreadsheet, only clear "+countrycode"
    driver.get("https://wa.me/"+countrycode+str(phone)+"?text=" + text)
    wait = WebDriverWait(driver, 30)

    time.sleep(15)
    x_arg = '//*[@id="action-button"]'
    continue_box = wait.until(EC.presence_of_element_located((
        By.XPATH, x_arg)))
    continue_box.click()
    time.sleep(7)
    web_text = '//*[@id="fallback_block"]/div/div/a'
    use_web = wait.until(EC.presence_of_element_located((
        By.XPATH, web_text)))
    use_web.click()
    time.sleep(7)
    actions = ActionChains(driver)
    actions.send_keys(Keys.ENTER).perform()
    time.sleep(10)


driver.close()
print("Mensajes enviados")