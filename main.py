from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import time 
import pandas as pd
import pandas 
import numpy
import openpyxl.styles.numbers
import smtplib
import mimetypes                                         
from email import encoders                                
from email.mime.base import MIMEBase                     
from email.mime.text import MIMEText 
from email.mime.multipart import MIMEMultipart                                       
import os
import datetime
from calendar import monthrange

browser = webdriver.Firefox(executable_path='C:/geckodriver/geckodriver.exe')
browser.maximize_window()     

browser.get("https://www.moex.com/") 
time.sleep(2)

#menu
browser.find_element(By.XPATH, "/html/body/div[3]/div[2]/div/div/div[2]/nav/span[1]/button").click() 
time.sleep(2)

#futures_market
button = browser.find_element(By.XPATH, '//*[@id="redesign-2021"]/div[3]/div[2]/div/div/div[2]/nav/span[1]/div/div/div/div[1]/div[3]/a')
button.click()
time.sleep(2)

#agreement
button_agree = browser.find_element(By.XPATH, '//*[@id="content_disclaimer"]/div/div/div/div[1]/div/a[1]')
button_agree.click()
time.sleep(2)

#indicative_courses
browser.find_element(By.XPATH, '//*[@id="ctl00_frmLeftMenuWrap"]/div/div/div/div[2]/div[13]/a').click()
time.sleep(2)

#GET VALUES: get_first_day, get_last_day, get_last_month, get_year
now = datetime.datetime.now()
get_first_day = 1

get_last_month = now.month
if get_last_month == 1:
    get_last_month = 12
    get_year = now.year - 1
else: 
    get_last_month = get_last_month - 1
    get_year = now.year

get_last_day = monthrange(now.year, get_last_month)[1]


#USD
usd = browser.find_element(By.XPATH, '//select[@id="ctl00_PageContent_CurrencySelect"]/option[@value="USD_RUB"]').click()

#first_day
select = Select(browser.find_element(By.ID, "d1day"))
select.select_by_value(str(get_first_day)) 

#month
select = Select(browser.find_element(By.ID, "d1month"))
select.select_by_value(str(get_last_month)) 

#year
select = Select(browser.find_element(By.ID, "d1year"))
select.select_by_value(str(get_year)) 


#last_day
select = Select(browser.find_element(By.ID, "d2day"))
select.select_by_value(str(get_last_day)) 

#month
select = Select(browser.find_element(By.ID, "d2month"))
select.select_by_value(str(get_last_month)) 

#year
select = Select(browser.find_element(By.ID, "d2year"))
select.select_by_value(str(get_year)) 

button=browser.find_element(By.XPATH, '/html/body/div[3]/div[3]/div/div/div[1]/div[2]/div/div/div/div[2]/form/div[4]/div[2]/div/div[5]/input').click()


#GET USD DATA TO TABLE

#get date column
date_usd = browser.find_elements(By.XPATH, '//table[@class="tablels"]/tbody/tr/td[1]')
#get value column
value_usd = browser.find_elements(By.XPATH, '//table[@class="tablels"]/tbody/tr/td[4]')
#get time column
time_usd = browser.find_elements(By.XPATH, '//table[@class="tablels"]/tbody/tr/td[5]')

df_building = pd.DataFrame(columns=['Дата USD/RUB', 'Курс USD/RUB', 'Время USD/RUB'])

for i in range(len(date_usd)):
    df_building = df_building.append({'Дата USD/RUB':date_usd[i].text, 'Курс USD/RUB': float((value_usd[i].text).replace(',', '.')), 'Время USD/RUB': time_usd[i].text}, ignore_index=True)
df_building.to_excel('Result1.xlsx', index=False)

#JPY
jpy = browser.find_element(By.XPATH, '//select[@id="ctl00_PageContent_CurrencySelect"]/option[@value="JPY_RUB"]').click()

#first_day
select = Select(browser.find_element(By.ID, "d1day"))
select.select_by_value(str(get_first_day)) 

#month
select = Select(browser.find_element(By.ID, "d1month"))
select.select_by_value(str(get_last_month)) 

#year
select = Select(browser.find_element(By.ID, "d1year"))
select.select_by_value(str(get_year)) 

#last_day
select = Select(browser.find_element(By.ID, "d2day"))
select.select_by_value(str(get_last_day)) 

#month
select = Select(browser.find_element(By.ID, "d2month"))
select.select_by_value(str(get_last_month)) 

#year
select = Select(browser.find_element(By.ID, "d2year"))
select.select_by_value(str(get_year)) 

#get date column
date_jpy = browser.find_elements(By.XPATH, '//table[@class="tablels"]/tbody/tr/td[1]')
#get value column
value_jpy = browser.find_elements(By.XPATH, '//table[@class="tablels"]/tbody/tr/td[4]')
#get time column
time_jpy = browser.find_elements(By.XPATH, '//table[@class="tablels"]/tbody/tr/td[5]')

building = pd.DataFrame(columns=['Дата JPY/RUB', 'Курс JPY/RUB', 'Время JPY/RUB'])

#Save to Excel 
for i in range(len(date_jpy)):
    building = building.append({'Дата JPY/RUB':date_jpy[i].text, 'Курс JPY/RUB': float((value_jpy[i].text).replace(',', '.')), 'Время JPY/RUB': time_jpy[i].text}, ignore_index=True)
building.to_excel('Result2.xlsx', index=False)

#merging_into_common_file
f1 = pandas.read_excel("Result1.xlsx")
f2 = pandas.read_excel("Result2.xlsx")
f3 = f1.join(f2)
f3.to_excel("Result.xlsx", index=False)

#add_column_Result
df = pd.read_excel("Result.xlsx")
df["Результат"] = df["Курс USD/RUB"]/df["Курс JPY/RUB"]
df.to_excel("Result.xlsx", index=False)

# Auto-fit_columns_width
writer = pd.ExcelWriter('Result.xlsx') 
df.to_excel(writer, sheet_name='Sheet1', index=False, na_rep='NaN')

for column in df:
    column_width = max(df[column].astype(str).map(len).max(), len(column))
    col_idx = df.columns.get_loc(column)
    writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)
writer.save()

df = pd.read_excel("Result.xlsx")
writer = pd.ExcelWriter("Result.xlsx", engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1', index=False)

#change_cells_format
workbook  = writer.book
worksheet = writer.sheets['Sheet1']
format1 = workbook.add_format({'num_format': '#,##0.0000_);(#,##0.0000)'})
worksheet.set_column('B:B', 10, format1)
worksheet.set_column('E:E', 10, format1)

#count_num_rows
file = pd.read_excel("Result.xlsx")
df = pd.DataFrame(file)
x = df.shape[0] + 1

#declination
if (x % 10 != 1 and x % 10 != 2 and x % 10 != 3 and x % 10 != 4) or x % 100 == 11 or x % 100 == 12 or x % 100 == 13 or x % 100 == 14:
    txt = "строк в документе Excel"
elif x % 10 == 2 or x % 10 == 3 or x % 10 == 4:
    txt = "строки в документе Excel"
else: 
    txt = "строка в документе Excel"

#send_email
addr_from = "induffer@gmail.com"                   
addr_to   = "induffer@gmail.com"                     
password  = "sjupcwqiohodqrzw"                                  

msg = MIMEMultipart()                               
msg['From']    = addr_from                          
msg['To']      = addr_to                            
msg['Subject'] = 'test.py'                   

body = str(x) + " " + txt
msg.attach(MIMEText(body, 'plain'))                 

#addFile
filepath="Result.xlsx"                   
filename = os.path.basename(filepath)                     
if os.path.isfile(filepath):                              
  ctype, encoding = mimetypes.guess_type(filepath)        
  maintype, subtype = ctype.split('/', 1)                 
  if maintype == 'text':                                  
      with open(filepath) as fp:                          
          file = MIMEText(fp.read(), _subtype=subtype)    
          fp.close()                                      
  else:                                                   
      with open(filepath, 'rb') as fp:
          file = MIMEBase(maintype, subtype)              
          file.set_payload(fp.read())                     
          fp.close()
      encoders.encode_base64(file)                        
  file.add_header('Content-Disposition', 'attachment', filename=filename) 
  msg.attach(file)                                        

server = smtplib.SMTP('smtp.gmail.com', 587)           
server.set_debuglevel(True)                         
server.starttls()                                   
server.login(addr_from, password)                   
server.send_message(msg)                            
server.quit()                                      

time.sleep(10)
browser.quit()






