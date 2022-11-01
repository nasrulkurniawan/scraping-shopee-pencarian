from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import pandas as pd
import time
import xlsxwriter

opsi = webdriver.ChromeOptions()
opsi.add_argument('--headless')
servis = Service('chromedriver.exe')
driver = webdriver.Chrome(service=servis, options=opsi)

shopee_link = "https://shopee.co.id/search?keyword=source%20code"
driver.set_window_size(1300,800)
driver.get(shopee_link)

rentang = 500
for i in range(1,7):
    akhir = rentang * i 
    perintah = "window.scrollTo(0,"+str(akhir)+")"
    driver.execute_script(perintah)
    print("loading ke-"+str(i))
    time.sleep(1)

time.sleep(5)
driver.save_screenshot("home.png")
content = driver.page_source
driver.quit()

data = BeautifulSoup(content,'html.parser')
# print(data.encode("utf-8"))

i = 0
base_url = "https://shopee.co.id"

list_number,list_nama,list_kota,list_gambar,list_harga,list_link,list_terjual,list_lokasi=[],[],[],[],[],[],[],[]

for area in data.find_all('div',class_="col-xs-2-4 shopee-search-item-result__item"):
    print('proses data ke-'+str(i))
    nama = area.find('div',class_="ie3A+n bM+7UW Cve6sh").get_text()
    kota = area.find('div',class_="zGGwiV").get_text()
    # gambar = area.find('img')['src']
    gambar = area.find('img').get('src')
    harga = area.find('span',class_="ZEgDH9").get_text()
    link = base_url + area.find('a')['href']
    terjual = area.find('div',class_="r6HknA uEPGHT")
    if terjual != None:
        terjual = terjual.get_text()
    lokasi = area.find('div',class_="zGGwiV").get_text()

      
    list_nama.append(nama)
    list_kota.append(kota)
    list_gambar.append(gambar)
    list_harga.append(harga)
    list_link.append(link)
    list_terjual.append(terjual)
    i+=1
    list_number.append(i)
    print("------")

df = pd.DataFrame({'Nomor':list_number,'Nama':list_nama,'Kota':list_kota,'Gambar':list_gambar,'Harga':list_harga,'Terjual':list_terjual,'Link':list_link})
with pd.ExcelWriter('shopee.xlsx') as writer:
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    border_fmt = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1})
    worksheet.conditional_format(xlsxwriter.utility.xl_range(0, 0, len(df), len(df.columns)), {'type': 'no_errors', 'format': border_fmt})
 