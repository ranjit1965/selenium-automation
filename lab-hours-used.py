from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

from openpyxl.workbook import Workbook
from openpyxl import load_workbook
import time



book_name="book.xlsx"       # mention name of the excel sheet
start_row = 2               # mention the starting row
end_row = 4                 # mention the ending row
rhn_col="D"                 # mention the col of the rhn id in excel sheet
lab_usd_col="AL"            # mention the col of the lab used in excel sheet
passwd="Itech@123"          # mention the password of rhn id's

wb=Workbook()
wb=load_workbook(book_name)
ws=wb.active


for row in range(start_row, end_row):
    
    rhn_cell = f"{rhn_col}{row}"
    data=ws[rhn_cell].value
    if data is None:
        continue
    
    if data == 'No' or data == 'NO' or data == 'no':
        continue
    
    driver = webdriver.Chrome()
    driver.maximize_window()
    
    driver.get("https://rha.ole.redhat.com/rha/app/login")
    
    time.sleep(5)
    
    email_elem = driver.find_element(By.NAME, "username") 
    email_elem.send_keys(data)
    cookie_btn = driver.find_element(By.ID,'truste-consent-button')
    cookie_btn.click()
    next_btn = driver.find_element(By.ID, "login-show-step2")  
    next_btn.click()
    #time.sleep(2)
    
    
    time.sleep(5)
    password_elem = driver.find_element(By.ID, "password")
    password_elem.send_keys(passwd)
    
    #time.sleep(5)
    
    login_box = driver.find_element(By.ID, 'rh-password-verification-submit-button')
    login_box.click()
    time.sleep(10)
    

    WebDriverWait(driver=driver, timeout=10).until(
    lambda x: x.execute_script("return document.readyState === 'complete'")
    )
    try:
        wait = WebDriverWait(driver, 7)
        element = wait.until(EC.visibility_of_element_located((By.ID, "rh-login-form-error-title")))
        if element is not  None :
            lab_used_cell=f"{lab_usd_col}{row}"
            ws[lab_used_cell]="credential error"
            driver.quit()
            continue
    except TimeoutException:
        pass
    
    #time.sleep(2)
    #driver.find_element(By.TAG_NAME,'body').send_keys(Keys.CONTROL, '-')

    
    launch_box = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'launch-button') and contains(@class, 'ml-2') and contains(@class, 'btn') and contains(@class, 'btn-link')]")))
    #driver.execute_script("arguments[2].scrollIntoView(true);", launch_box)
    
    
    launch_box.click()
    time.sleep(5)


    lab_box = driver.find_element(By.ID, 'course-tabs-tab-8')
    lab_box.click()
    time.sleep(5)
    
    lab_hour = driver.find_element(By.CLASS_NAME, 'instruction-wrapper')
    
    used=str(lab_hour.text)
    lab_used_cell=f"{lab_usd_col}{row}"
    ws[lab_used_cell]=used[16:]
    
    driver.quit()


wb.save(book_name)
    

