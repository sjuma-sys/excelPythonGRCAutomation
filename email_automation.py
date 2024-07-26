import win32com.client
import datetime
from time import sleep


dispatcher = win32com.client.Dispatch('Outlook Application')

mailsize = 0x0

newmail = dispatcher.CreateItem(mailsize)

newmail.To = input('Who will be recieveing the email, seperate multiple with commas please')

newmail.CC = input('Who will be CCd in the email, seperate with commas please')
newmail.CC = newmail.CC.replace(",",";")

newmail.body = input('What will be contained in the body of the email')

newmail.BCC = input('Who will be CCd in the email, seperate with commas please')

date = datetime.date.today() + datetime.timedelta(days=15)	#adds on additional days to todays date

date = date.strftime("%d,%m,%y")	#changing the format of said date

newmail.HTMLbody += "<HTML> </HTML>"	#changes the body to HTMLbody can be appended to like a string with the +=

newmail.Display()	#display the mail

#newmail.send()	#sends the mail