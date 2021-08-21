from BsClass import *

username = ""
password = ""
drive = Driver(username=username,password=password)

drive.LoginCO()
AllList = drive.AllList()
TextList = drive.TextList(AllList)
DateList = drive.DateList(AllList)
MixedList = drive.MixedList(DateList,TextList)


FirstSentence = ""
To = ""
Subject = ""
CC = ""
drive.MixedListSendMail(MixedList,FirstSentence,To,Subject,CC)
