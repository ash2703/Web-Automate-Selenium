from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from openpyxl import Workbook  #for working on excel sheets
from time import sleep

# Using Chrome to access web
# options = webdriver.ChromeOptions()
# options.add_argument("--start-maximized")
# driver = webdriver.Chrome("D:\Softwares\chromedriver", chrome_options=options)   #path to chrome driver
driver = webdriver.Chrome("D:\Softwares\chromedriver")
wait = WebDriverWait(driver, 10)


webpage =  'https://internshala.com/'
username = "your username"
password = "your password"

path = "E:\Codes\Python\Web-Automate-Selenium\Internshala-Python-Users.xlsx"


def login(username,password):
    driver.find_element_by_xpath("//div[@id = 'register-button-positioner']//button[@type = 'button']").click()

    user = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='modal_email']"))).send_keys(username)

    Pass = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='modal_password']"))).send_keys(password)
    
    driver.find_element_by_xpath("//*[@id='modal_login_submit']").click()

def scrollDown(driver):
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Scroll down the page
def scrollDownAllTheWay(driver):
    old_page = driver.page_source
    while True:
        for i in range(2):
            scrollDown(driver)
            sleep(2)
        new_page = driver.page_source
        if new_page != old_page:
            old_page = new_page
        else:
            break
    return True

# Open the website
driver.get(webpage)
login(username,password)

wait.until(EC.url_to_be("https://internshala.com/internships/matching-preferences"))   #same element was present on previous page so error was thrown
menu = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='dropdown']/ul/li[2]/a"))).click()

window_after = driver.window_handles[1]   #focus on newly opened window
driver.switch_to.window(window_after)   #switch to the new tab to access its elements
 
 
dash = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="dropdown"]/ul/li[1]/a'))).click()   #Dasboard

course = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="trainings_container"]/ul/li[2]/p[5]/a[2]'))).click()  #Course manager

project = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="project_evaluation_menu_item"]/p'))).click()   #Evaluate Project

contest_user = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="project_evaluation"]/div/div[1]/label[1]')))  #Uncheck contest users

scrollDownAllTheWay(driver)  #Scroll in order to load the whole table

print(contest_user.is_selected())  #check if box already pressed

# if not contest_user.is_selected():  #get_attribute("checked") return none if not selected
#     contest_user.click()  #internshala page is checked when false is returned

table = driver.find_element_by_xpath('//*[@id="project_evaluation"]/div/div[2]/div/div[1]/table')   #Read each row of table

rows = table.find_elements_by_xpath(".//tr")   #fetch all rows from the table
print(len(rows))   #total no. of students

book = Workbook()   #create an excel sheet
nonContestUsers = book.create_sheet("non_contest_users",0)  #create a sheet for non-contest users
contestUsers = book.create_sheet("contest_users",1)  #create a sheet for contest users
contestUsers.append(("Name", "Start date", "Submission date", "Download link"))
nonContestUsers.append(("Name", "Start date", "Submission date", "Download link"))

contestUsersNames = set()   #set to add only unique names in the sheet
nonContestUsersNames = set()

for row in rows[1:]:
    name = row.find_element_by_xpath(".//td[1]").text    # .//tr/td[1]    name
    if len(name) > 2 and (name not in contestUsersNames and name not in nonContestUsersNames):
        start_date = row.find_element_by_xpath(".//td[2]").text        # .//tr/td[2]    start date
        submission_date = row.find_element_by_xpath(".//td[3]").text   # .//tr/td[3]    submission date
        download_link = row.find_element_by_xpath(".//td[4]//a").get_attribute("href")  # .//tr/td[4]    download link
                                                                                        # .//tr/td[5]    share feedback button
        if "\n" in name:
            contestUsersNames.add(name.split("\n")[0])
            contestUsers.append((name.split("\n")[0], start_date, submission_date, download_link))
        else:
            nonContestUsersNames.add(name)
            nonContestUsers.append((name, start_date, submission_date, download_link))

book.save(path)
print("Done Saving")