#Script crafted by Ryan Siegel for NOC use. This script will be started by task scheduler daily and then prints to the label maker the tape swap info.
import asyncio,os,win32print
import win32ui
from datetime import datetime, timedelta, date

#change between week 1 and 2
def week1or2(currentDate):
    yearWeek = currentDate.isocalendar()[1]
    numDay = currentDate.weekday()
    if yearWeek % 2 == 0:
        currentWeek = "1"
        if numDay == 6:
            currentWeek = "2"
    else:
        currentWeek = "2"
        if numDay == 6:
            currentWeek = "1"
    return currentWeek

#check which day of the week it is
def dayOfWeek(currentDate):
    numDay = currentDate.weekday()
    if numDay == 0:
        return "Mon"
    elif numDay == 1:
        return "Tues"
    elif numDay == 2:
        return "Wed"
    elif numDay == 3:
        return "Thurs"
    elif numDay == 4:
        return "Fri"
    elif numDay == 5:
        return "Sat"
    elif numDay == 6:
        return "Sun"

#check which day of the week it is
def currentMonth(account, currentDate, accountOne):
    numMonth = currentDate.month
    if numMonth == 1:
        if account == accountOne:
            return "Dec"
        else:
            return "Jan"
    elif numMonth == 2:
        if account == accountOne:
            return "Jan"
        else:
            return "Feb"
    elif numMonth == 3:
        if account == accountOne:
            return "Feb"
        else:
            return "Mar"
    elif numMonth == 4:
        if account == accountOne:
            return "Mar"
        else:
            return "Apr"
    elif numMonth == 5:
        if account == accountOne:
            return "Apr"
        else:
            return "May"
    elif numMonth == 6:
        if account == accountOne:
            return "May"
        else:
            return "June"
    elif numMonth == 7:
        if account == accountOne:
            return "June"
        else:
            return "July"
    elif numMonth == 8:
        if account == accountOne:
            return "July"
        else:
            return "Aug"
    elif numMonth == 9:
        if account == accountOne:
            return "Aug"
        else:
            return "Sept"
    elif numMonth == 10:
        if account == accountOne:
            return "Sept"
        else:
            return "Oct"
    elif numMonth == 11:
        if account == accountOne:
            return "Oct"
        else:
            return "Nov"
    elif numMonth == 12:
        if account == accountOne:
            return "Nov"
        else:
            return "Dec"

#check if last day of month
def lastDayOfMonth(currentDate):
    tomorrowsDate = currentDate + timedelta(days=1)
    if tomorrowsDate.month != currentDate.month:
        return True
    else:
        return False

#check if first sat of month
def firstSatOfMonth(currentDate):
    if currentDate.weekday() == 5:
        if currentDate.day >= 1 and currentDate.day <= 7:
            return True
        else:
            return False
    else:
        return False

def wedOrSat(currentDate):
    if currentDate.weekday() == 2 or currentDate.weekday() == 5:
        return True
    else:
        return False

accountOne = "ACCOUNT ONE NAME HERE"
accountTwo = "ACCOUNT TWO NAME HERE"
accountThree = "ACCOUNT THREE NAME HERE"
accountFour = "ACCOUNT FOUR NAME HERE"
labelText = []
currentDate = datetime.now() - timedelta(days=1) #grabs todays date then subtracts one
currentWeek = week1or2(currentDate) #determines if it is week 1 or 2
dayWeek = dayOfWeek(currentDate) #determines day of week
accountOneMonthly = firstSatOfMonth(currentDate) #determines if first sat of the month
accountTwoMonthly = lastDayOfMonth(currentDate) #determines if last day of the month

#//////Account One
#check if 2 tapes or 1
accountOne2 = wedOrSat(currentDate)
if accountOne2 == True:
    labelText.append(accountOne + " LTO | " + dayWeek + " Daily - " + str(currentDate.month) + "/" + str(currentDate.day) + " - Week " + currentWeek + " | 2 Tapes\n")
else:
    labelText.append(accountOne + " LTO | " + dayWeek + " Daily - " + str(currentDate.month) + "/" + str(currentDate.day) + " - Week " + currentWeek + " | 1 Tape\n")
#check to add monthly tape
if accountOneMonthly == True:
    account = accountOne
    monthName = currentMonth(account, currentDate, accountOne) #determines last month
    labelText.append(accountOne + " LTO | " + monthName + " Monthly - " + str(currentDate.month) + "/" + str(currentDate.day) + " | 2 Tapes")

#//////Account Two
labelText.append(accountTwo + " LTO | " + dayWeek + " Daily - " + str(currentDate.month) + "/" + str(currentDate.day) + " - Week " + currentWeek + " | 1 Tape")
#check to add monthly tape
if accountTwoMonthly == True:
    account = accountTwo
    monthName = currentMonth(account, currentDate, accountOne) #determines current month
    labelText.append(accountTwo + " LTO | " + monthName + " Monthly - " + str(currentDate.month) + "/" + str(currentDate.day) + " | 1 Tape")

#//////Account Three
labelText.append(accountThree + " LTO | " + dayWeek + " Daily - " + str(currentDate.month) + "/" + str(currentDate.day) + " - Week " + currentWeek + " | 1 Tape\n")

#//////Account Four
labelText.append(accountFour + " LTO | " + dayWeek + " Daily - " + str(currentDate.month) + "/" + str(currentDate.day) + " - Week " + currentWeek + " | 1 Tape")

#Print label
X=50; Y=35 #Position on the label
dc = win32ui.CreateDC()
dc.CreatePrinterDC("PRINTER NAME HERE") #put printer name here
dc.StartDoc('Test document') #test to make sure printer is working, will print spoof page
dc.StartPage() #actually print label
#set font
fontdata = {'height':35} #size
font = win32ui.CreateFont(fontdata)
dc.SelectObject(font)
for line in labelText:
     dc.TextOut(X,Y,line)
     Y += 40 #make 5 above the the size, this adds to the position on the label
#close printing
dc.EndPage()
dc.EndDoc()