from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

drivepath="C:\\Users\MYPC\Desktop\chromedev\chromedriver"
newdrive=webdriver.Chrome(drivepath)
newdrive.get('https://www.flipkart.com/search?q=airpods+pro&sid=0pm%2Cfcn%2C821%2Ca7x%2C2si&as=on&as-show=on&otracker=AS_QueryStore_OrganicAutoSuggest_2_3_na_na_na&otracker1=AS_QueryStore_OrganicAutoSuggest_2_3_na_na_na&as-pos=2&as-type=RECENT&suggestionId=airpods+pro%7CTrue+Wireless&requestId=4247eca9-a7d5-4f9d-bf74-1718b7acb8a2&as-searchtext=air')
titles=[]
amounts=[]
books=newdrive.find_elements(by=By.CLASS_NAME,value="s1Q9rs")
for book in books:
    titles.append(book.text)
prices=newdrive.find_elements(by=By.CLASS_NAME,value="_30jeq3")
for price in prices:
    amounts.append(price.text)
newbook=Workbook()
sheet=newbook.active
sheet["A1"]="Booktitle"
sheet["B1"]="Price"

for i in range(len(titles)):
    sheet[f"A{i+2}"]=titles[i]

for i in range(len(amounts)):
    sheet[f"B{i+2}"]=amounts[i]
newbook.save("Products.xlsx")


