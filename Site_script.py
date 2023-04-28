import requests
import re
import xlwt
from bs4 import BeautifulSoup
from time import sleep
import openpyxl

wBook = openpyxl.load_workbook('mydata.xlsx')
sheet = wBook.active
wb = xlwt.Workbook()
ws = wb.add_sheet('sheet1')
def url_input(url):
    Data = []
    print("Create Data List for collect Data..."),print(),print(),print()
    print("-*"*16)
    #connect to website...
    print("connect to web site...")
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")        
        title = soup.find('body').text 
        print("connect to web server successfull"),print(),print(),print()
        print("-*"*14)
        #extract company name code...
        print("Extracting Data For company_name.."),print(),print(),print()
        print("-*"*10)
        company_name = soup.find('h4', attrs = {'class': 'text-right'}).text
        Data.append(company_name)
        print("company name data is collected successfull..."),print(),print(),print()
        print("-*"*10)
        
        #Find and Extract Website Address
        print("Extract data for web site address...."),print(),print(),print()
        print("-"*29)
        for link in soup.find_all('a', attrs = {'class': 'btn btn-primary shadow-btn hover-ads'}):
            site_address = link.get("href")
        Data.append(site_address)
        print("web site data collect successfull..."),print(),print(),print()
        print("-"*32)
        #extract phone Data...
        print("Extacting Data For address..."),print(),print(),print()
        print("-*"*10)
        address = soup.find_all('p', attrs = {'class': 'text-justify rtl'})
        t = tuple(address)
        x = str(t[2])
        #Extract contant with regex...
        print("rexex is starting....")
        result =re.findall("<p.+\s*.*p>",x)[0]
        result = str(re.sub('<.*?>', '', result)) 
        address = result.strip()
        Data.append(address)
        print("address data collect successfull...")
        print("-"*19)        
        print("Extracting data for phone , mobile and fax"),print(),print(),print()
        print("-"*32)
        mobile_fax = soup.find_all('p', attrs = {'class': 'text-right ltr'})
        phone_list = []
        mobile_list = []
        fax_list = []
        for mob in mobile_fax:
            if '</span> <span class="bolding">:تلفن</span>' in str(mob):
                phone = str(mob).replace('<p class="text-right ltr">',"").replace('</span> <span class="bolding">:تلفن</span>',"").replace('</p>',"").replace('<span>',"").split("،")
                for pho in phone:
                    phone_numbers = pho.strip()
                    phone_list.append(phone_numbers)
            if  '</span> <span class="bolding">:همراه</span>' in str(mob):
                mobile = str(mob).replace('<p class="text-right ltr">',"").replace('</span> <span class="bolding">:همراه</span>',"").replace('<span>',"").replace('</p>',"").split("،")
                print("Extract Data mobile...")
                for mob in mobile:
                    mobile = mob.strip()
                    mobile_list.append(mobile)
            if  '<span class="bolding">:فاکس</span>' in str(mob):
                fax = str(mob).replace('<p class="text-right ltr">',"").replace('</span> <span class="bolding">:فاکس</span>',"").replace('</p>',"").replace('<span>',"").split("،")
                for fa in fax:
                    fax = fa.strip()
                    fax_list.append(fax) 
        print("extracting data successfull")
        print("-*"*32)
        #Find address html tag and content...
        print(),print(),print()
        print(f"extracting data {company_name}"),print(),print(),print()
    except:
        print("""Warning!!
                 Error Data Collect Level....""")
    return Data,phone_list,mobile_list,fax_list

ws.write(0, 3, 'نام شرکت')
ws.write(0, 4, 'سایت')
ws.write(0, 5, 'ادرس')
counter =1
def save_data(Data,phone_list,mobile_list,fax_list):
    print("Starting Save Data on Excel")
    global counter
    print(),print(),print()
    print("start add data to excel.")
    print()
    print("add company name.")
    print()
    try:
        ws.write(counter, 3,Data[0])
        ws.write(counter, 4,Data[1])
        print("add web site address..")
        print()
        ws.write(counter, 5,Data[2])
        if len(phone_list) !=0:
            counter_phone = 7
            for i in phone_list:
                ws.write(counter,counter_phone,i)
                counter_phone +=1

        if len(mobile_list) !=0:
            counter_mobile =10
            for i in mobile_list:
                ws.write(counter,counter_mobile,i)
                counter_mobile +=1

        if len(fax_list) !=0:
            counter_fax = 13
            for i in fax_list:
                ws.write(counter,counter_fax,i)
                counter_fax +=1
        print("saveing Data to excel..")
        wb.save('test.xls')
        print("save Data Successfull"),print(),print(),print()
        counter +=1
        print("Done......")
    except:
        print("""Warning You Have One Error!
                 Your Error:
                 Excel Is open please close The excel file Please....""")
while True:
   try:
     usr_url= input("please Enter url address: ")
     Data,phone_list,mobile_list,fax_list = url_input(usr_url)
     save_data(Data,phone_list,mobile_list,fax_list)
   except:
       print("""Warning Error please Try Again Enter your url address
                 your error maby:
                 Connection Error
                 Url Not Found""")
