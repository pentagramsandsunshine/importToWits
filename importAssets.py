
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
import passwords_filelocations as pf

#Set up selenium
options = webdriver.ChromeOptions()
options.add_experimental_option("useAutomationExtension", False)
options.add_experimental_option("excludeSwitches",["enable-automation"])
#desktop
options.add_argument(r"user-data-dir=C:\Users\lukek\AppData\Local\Google\Chrome\User Data\Profile 1")
driver_path = r"C:\Program Files (x86)\chrome-win64\chromedriver.exe"
driver = webdriver.Chrome(driver_path, options=options)

def loginInWits(username, password):
    wait = WebDriverWait(driver, 5)
    try:
        driver.get('https://wits.itrenew.com/')
        wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div/div/div/div/form/div[2]/div[1]/input')))
        try:
            #Check if username and password saved using autofill
            driver.find_element(By.XPATH, '/html/body/div/div/div/div/form/div[2]/button').click()
            wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/form/div[3]/nav/div/div[2]/div/ul[1]/li[4]/a')))
            print('username and password saved')
        except:
            #log in using given username / password
            usernameField = driver.find_element(By.XPATH, '/html/body/div/div/div/div/form/div[2]/div[1]/input')
            usernameField.send_keys(Keys.CONTROL + 'a')
            time.sleep(1)
            usernameField.send_keys(Keys.BACK_SPACE)
            time.sleep(.2)
            usernameField.send_keys(username)

            passwordField = driver.find_element(By.XPATH, '/html/body/div/div/div/div/form/div[2]/div[2]/input')
            passwordField.send_keys(Keys.CONTROL + 'a')
            time.sleep(1)
            passwordField.send_keys(Keys.BACK_SPACE)
            time.sleep(.2)
            passwordField.send_keys(password)
            time.sleep(.1)
            #Click log in
            driver.find_element(By.XPATH, '/html/body/div/div/div/div/form/div[2]/button').click()
    except:
        print('issues')
        time.sleep(5)

def bulkImportExcel(fileLoc, job):
    loginInWits(pf.witsUser, pf.witsPass)
    wait = WebDriverWait(driver, 15)

    #nav to import
    wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/form/div[3]/nav/div/div[2]/div/ul[1]/li[4]/a')))
    time.sleep(.1)
    dropDown = driver.find_element(By.XPATH, '/html/body/form/div[3]/nav/div/div[2]/div/ul[1]/li[4]/a')
    dropDown.click()
    wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/form/div[3]/nav/div/div[2]/div/ul[1]/li[4]/ul/li[5]/a')))
    importPage = driver.find_element(By.XPATH, '/html/body/form/div[3]/nav/div/div[2]/div/ul[1]/li[4]/ul/li[5]/a')
    importPage.click()
    wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/form/div[4]/main/div/div[3]/div/div/div/input')))
    jobInput = driver.find_element(By.XPATH, '/html/body/form/div[4]/main/div/div[3]/div/div/div/input')
    jobInput.send_keys(job)
    time.sleep(.1)
    driver.find_element(By.XPATH, '//*[@id="MainContent_JobButton"]').click()

    #import file
    time.sleep(5)
    driver.find_element(By.XPATH, '/html/body/div[2]/main/div[2]/form/div[1]/div/div/input').send_keys(fileLoc)
    time.sleep(1)

    driver.find_element(By.XPATH, '/html/body/div[2]/main/div[2]/form/button').click()
    wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/main/div/div/div[1]/div[1]/button')))
    time.sleep(1)
    print('importing')
    driver.find_element(By.XPATH, '/html/body/main/div/div/div[1]/div[1]/button').click()


bulkImportExcel(r"C:\Users\lukek\Desktop\Work\WorkBots\WorkBot11_5\Test.xlsx", "I2238079")

# /html/body/form/div[3]/nav/div/div[2]/div/ul[1]/li[4]/a drop down
#/html/body/form/div[3]/nav/div/div[2]/div/ul[1]/li[4]/ul/li[5]/a import