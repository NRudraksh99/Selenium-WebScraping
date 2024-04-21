from selenium import webdriver as wd
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from os.path import dirname as dr
import openpyxl as xl
p=dr(__file__)

u=input("Enter the URL: ")
name=input("Enter the name of the resultant file required: ")
service=Service(executable_path=f"{p}\\Chromedriver\\chromedriver.exe")
driver=wd.Chrome(service=service)

driver.get(u)
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
text=[i.text.strip() for i in driver.find_elements(By.XPATH,"//*[not(self::script) and not(self::style)]") if i.text.strip()]
images=[i.get_attribute("src") for i in driver.find_elements(By.TAG_NAME,"img")]
links=[i.get_attribute("href") for i in driver.find_elements(By.TAG_NAME,"a")]

wb=xl.Workbook()
result=wb.active
result["A1"],result["B1"],result["C1"] = "Text","Image URLs","Links"
for i,j in enumerate(text,start=2):
    result[f"A{i}"]=j
for i,j in enumerate(images,start=2):
   result[f"B{i}"]=j
for i,j in enumerate(links,start=2):
    result[f"C{i}"]=j
wb.save(f"{p}\\Result_Sheets\\{name}.xlsx")
print("Result has been successfully saved in the Result_Sheets directory!!")
driver.quit()