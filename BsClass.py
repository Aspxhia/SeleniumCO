from selenium import webdriver
import time
from bs4 import BeautifulSoup
from datetime import datetime
import win32com.client

class Driver(object):
    def __init__(self,username,password):
        self.username = username
        self.password = password


    def BuildDriver(self):
        self.browser = webdriver.Chrome(executable_path="chromedriver.exe")


    def LoginCO(self):
        self.BuildDriver()
        self.url = "https://www.mechsoft.com.tr/web?debug#view_type=list&model=account.analytic.line&menu_id=144"
        self.browser.get(self.url)
        self.browser.find_element_by_xpath("""//*[@id="login"]""").send_keys(self.username)
        self.browser.find_element_by_xpath("""//*[@id="password"]""").send_keys(self.password)
        self.browser.find_element_by_xpath("""//*[@id="wrapwrap"]/main/div/main/div/form/div[5]/button""").click()
        time.sleep(3)


    def AllList(self):
        AllList = list()
        content = self.browser.page_source
        soup = BeautifulSoup(content,"html.parser")
        for div in soup.find_all("td", {"class": "o_data_cell o_required_modifier"}):
            AllList.append(div.text)
        return AllList

    def DateList(self,AllList):
        DateList = list()
        AllListLen = int(len(AllList)/5)

        # ---------------------------------------------
        sayTarih = 0
        aralikTarihler = 0

        if sayTarih < AllListLen + 1:
            for x in range(1, AllListLen + 1):
                date = datetime.strptime(str(AllList[aralikTarihler]), '%d-%m-%Y')
                dt = date.strftime("%A")
                if dt == "Monday":
                    dt = "Pazartesi"
                elif dt == "Tuesday":
                    dt = "Salı"
                elif dt == "Wednesday":
                    dt = "Çarşamba"
                elif dt == "Thursday":
                    dt = "Perşembe"
                elif dt == "Friday":
                    dt = "Cuma"
                elif dt == "Saturday":
                    dt = "Cumartesi"
                else:
                    dt = "Pazar"

                DateList.append(dt)
                sayTarih += 1
                aralikTarihler += 5

        return DateList

    def TextList(self,AllList):

        TextList = list()
        AllListLen = int(len(AllList) / 5)
        sayTextler = 0
        aralikTextler = 2
        if sayTextler < AllListLen + 1:
            for y in range(1, AllListLen + 1):
                TextList.append(AllList[aralikTextler])
                sayTextler += 1
                aralikTextler += 5

        return TextList

    def MixedList(self,DateList,TextList):
        MixedList = list(zip(DateList, TextList))
        return MixedList

    def Exit(self,second):
        time.sleep(second)
        self.browser.close()

    def SendMail(self,To,Subject,Body,CC):
        outlook = win32com.client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = To
        mail.Subject = Subject
        mail.HTMLBody = Body
        mail.CC = CC
        mail.Send()
        print("Mail gönderildi")

    def MixedListSendMail(self,MixedList,FirstSentence,To,Subject,CC):
        allItems = "<HTML><BODY>"+FirstSentence+"<br><br>"
        for item in MixedList:
            allItems = allItems + str(item[0]) + "-" + str(item[1]) + "<br>"

        allItems = allItems + "<br><br> İyi çalışmalar dilerim,<br>Saygılarımla</BODY></HTML>"
        self.Exit(2)
        self.SendMail(To, Subject, allItems,CC)


