from time import sleep
# excel
import xlwt
from xlwt import Workbook
# selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
# move mouse
import pyautogui




PATH = r"C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe"

webdr = webdriver.Chrome(PATH)

  
# open website
webdr.get("https://www.mehe.gov.lb:83/")

wait = WebDriverWait(webdr, 600)

print('next')

holder_nb = wait.until(EC.presence_of_element_located((By.ID,'nbr')))
holder_nb.send_keys('1')


student_section = wait.until(EC.presence_of_element_located((By.ID,'select2-study-container'))) 
student_section.click()

# input your class

sleep(10)

search_btn = wait.until(EC.presence_of_element_located((By.ID,'btnSearch')))
search_btn.click()

# print(f'biology: {grades[39] + grades[40]}')
# print(f'chemistry: {grades[54] + grades[55]}')
# print(f'physics: {grades[69] + grades[70]}')
# print(f'Math: {grades[85] + grades[86]}')
# print(f'arabic language: {grades[107] + grades[108]}')
# print(f'english language: {grades[130] + grades[131]}')


# Workbook is created
wb = Workbook()

# add_sheet is used to create sheet.
LS = wb.add_sheet('LS')

x = 1

LS.write(0, 0, 'رقم الطالب')
LS.write(0, 1, 'عام')
LS.write(0, 2, 'الدورة')
LS.write(0, 3, 'علوم الحياة')
LS.write(0, 4, 'كيمياء')
LS.write(0, 5, 'فيزياء')
LS.write(0, 6, 'رياضيات')
LS.write(0, 7, 'اللغة العربية')
LS.write(0, 8, 'اللغة الأجنبية')

wb.save('xlwt example.xls')


while True:

    grades = wait.until(EC.presence_of_element_located((By.ID,'student_notes'))).text

    LS.write(x, 0, x)
    LS.write(x, 1, 2022)
    LS.write(x, 2, 1)
    LS.write(x, 3, grades[39] + grades[40])
    LS.write(x, 4, grades[54] + grades[55])
    LS.write(x, 5, grades[69] + grades[70])
    LS.write(x, 6, grades[85] + grades[86])
    LS.write(x, 7, grades[107] + grades[108])
    LS.write(x, 8, grades[130] + grades[131])

    wb.save('xlwt example.xls')

    holder_nb.clear()
    holder_nb.send_keys(x)

    search_btn.click()

    x += 1   


    
