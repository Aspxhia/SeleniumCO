from BsClass import *

username = "tugkan.ozkan@mechsoft.com.tr"
password = "5943b77D."
drive = Driver(username=username,password=password)

drive.LoginCO()
AllList = drive.AllList()
TextList = drive.TextList(AllList)
DateList = drive.DateList(AllList)
MixedList = drive.MixedList(DateList,TextList)


FirstSentence = "Merhabalar Kerim Bey,<br>    Haftalık çalışma notlarım şu şekildedir:"
To = ""
Subject = ""
CC = ""
drive.MixedListSendMail(MixedList,FirstSentence,To,Subject,CC)
