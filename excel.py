from urllib.request import urlopen as uReq
import requests
from bs4 import BeautifulSoup
import time
import xlsxwriter

wbook = xlsxwriter.Workbook('d:\\project\\python\\togel\\hongkong1.xlsx')
wsheet = wbook.add_worksheet("Data")

url = "https://keluarannomor.org"

r = 3
response = requests.get(url)
if response.status_code==200:
    for i in range(1,107):
        url="https://keluarannomor.org/hasil-keluaran-togel-hongkong/"+str(i)
        resp = requests.get(url)
        if resp.status_code==200:
            uClient = uReq(url)
            halaman = uClient.read()
            uClient.close()
            soup = BeautifulSoup(halaman,"html.parser")
            #print(soup.title)
            tbl = soup.findAll("table")
            body = tbl[0].tbody
            trs = body.findAll("tr")
            for tr in trs:
                kolom = tr.findAll("td")
                n = kolom[2].get_text()
                
                wsheet.write("A"+str(r),kolom[0].get_text())
                wsheet.write("B"+str(r),kolom[1].get_text())
                wsheet.write("C"+str(r),int(n[0]))
                wsheet.write("D"+str(r),int(n[1]))
                wsheet.write("E"+str(r),int(n[2]))
                wsheet.write("F"+str(r),int(n[3]))
                r=r+1
            #print("\n\n")
        print("Dari halaman "+str(i)+" telah selesai!")
        r=r+1
    print("Excel telah terisi!")
    wbook.close()
    exit(1)
else:
    wbook.close()
    print("Koneksi internet putus!")
    exit(1)
