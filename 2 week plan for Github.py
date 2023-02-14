#The purpose of this script is to gather all work orders for the current day until 2 weeks from today
#and add them to an excel sheet. It will then create subtotals for each days planned routine maintenance hours
#and total planned hours for those 2 weeks
#Ishmam Raza Dewan
###############################COMMON PART#####################################

from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import datetime, time
import xlwings as xw
browser = webdriver.Chrome()

wb = xw.Book(r'INSERT')
sheet = xw.sheets[0]

#Changing what this script clicks on requires your browser dev tools
#Each object has an ID, inspect element to hover over object, click to find ID
#If planned hours = 0 in excel, change to 0.5. Make sure to use right excel book
browser.get('http://nscandacmaxapp1/maximo/webclient/login/login.jsp?welcome=true')
UserElem = browser.find_element(By.ID, "username")
UserElem.send_keys('INSERT HERE') #put your username here
passElem = browser.find_element(By.ID, "password")
passElem.send_keys('INSERT HERE') #put your password here
loginElem = browser.find_element(By.ID, "loginbutton")
loginElem.submit()
time.sleep(2)
wotrackElem = browser.find_element(By.ID, "FavoriteApp_WOTRACK") #make sure this is in your faves. To do a task not done through work order tracking, change this to a different menu
type(wotrackElem)
wotrackElem.click()
time.sleep(3)
###############################################################################
dt = datetime.datetime.now()
delta = datetime.timedelta(days = 14)
ft = dt + delta

smonth = str(dt.month)
sday = str(dt.day)
syear = str(dt.year)

fmonth = str(ft.month)
fday = str(ft.day)
fyear = str(ft.year)
time.sleep(1)
allElem = browser.find_element(By.ID, "m9e1854a7_ns_menu_queryMenuItem_0_a")
allElem.click()
time.sleep(4)
searchElem = browser.find_element(By.ID, "m68d8715f-tbb_text")
searchElem.click()
time.sleep(2)

istaskElem = browser.find_element(By.ID, "maa922243-tb")
istaskElem.send_keys('N')
time.sleep(1)

historyElem = browser.find_element(By.ID, "mdd9512d5-tb")
historyElem.send_keys('N')
time.sleep(1)
typeElem = browser.find_element(By.ID, "med325893-tb")
typeElem.send_keys('=HKG,=INR,=PPM')
time.sleep(3)

statusElem = browser.find_element(By.ID, "m449c436f-tb")
statusElem.send_keys('=RELEASED,=WPLAN,=WSCHED')
time.sleep(2)


startElem = browser.find_element(By.ID, "m3cdc438b-tb")
startElem.send_keys(smonth + '/' + sday + '/' + syear + ' 12:00 AM')
time.sleep(2)
finElem = browser.find_element(By.ID, "mac635e1a-tb")
finElem.send_keys(fmonth +'/' + fday + '/' +fyear + ' 12:00 AM')
time.sleep(2)


findElem = browser.find_element(By.ID, "m4fd840b0-pb")
findElem.click()
time.sleep(3)
numberofWOs = browser.find_element(By.ID, "m6a7dfd2f-lb3")
numberofWOs = numberofWOs.text
print(numberofWOs)
try:
    numberofWOs = numberofWOs[13] + numberofWOs[14] + numberofWOs[15]
except:
    numberofWOs = numberofWOs[11] + numberofWOs[12]
numberofWOs = int(numberofWOs)


time.sleep(2)
wo1Elem = browser.find_element(By.ID, "m6a7dfd2f_tdrow_[C:1]_ttxt-lb[R:0]") #change ID here if diff
type(wo1Elem) #wo1 = work order one
time.sleep(1)
wo1Elem.click()
time.sleep(1)
i=0
while i<numberofWOs:
    time.sleep(3)
    wb.save()
    WO = browser.find_element(By.ID, "m52945e17-tb").get_attribute("value");
    description = browser.find_element(By.ID, "md42b94ac-tb").get_attribute("value");
    jobplan = browser.find_element(By.ID, "mfe7bb84-tb").get_attribute("value");
    tstart = browser.find_element(By.ID, "m651c06b0-tb").get_attribute("value");
    asset = browser.find_element(By.ID, "m3b6a207f-tb").get_attribute("value");
    PlansElem = browser.find_element(By.ID, "m356798d1-tab_anchor")
    time.sleep(3)
    type(PlansElem)
    PlansElem.click()
    time.sleep(4)
    try:
        plannedhrs1 = browser.find_element(By.ID, "m5e4b62f0_tdrow_[C:9]_txt-tb[R:0]").get_attribute("value");
    except:
        plannedhrs1 = 'no data'
    try:
        plannedhrs = int(plannedhrs1[0])
    except:
        print('no data')

    try:
        plannedhrs = int(plannedhrs1[0] + plannedhrs1[1])
    except:
        print('no double digits or still no data')
        
    if plannedhrs1 == 'no data':
        plannedhrs = 'no data'
        
#from the above, if no data in plannedhrs1, then plannedhrs is no data. otherwise, it will test with 1
#character, and then test again with 2.

#plannedhrs 2 section
        
    try:
        plannedhrstwo = browser.find_element(By.ID, "m5e4b62f0_tdrow_[C:9]_txt-tb[R:1]").get_attribute("value");
    except:
        plannedhrstwo = 'no data'
    else:
        plannedhrs2 = plannedhrstwo[0]

    try:
        plannedhrs2 = int(plannedhrs2)
    except:
        print('no data')

    try:
        plannedhrs2 = int(plannedhrstwo[0] + plannedhrstwo[1])
    except:
        print('less than 10 hours or still no data')

        
    try:
        plannedhrs = plannedhrs + plannedhrs2
    except:
        print('no data for plannedhrs2')


#plannedhrs 3 section (identical to plannedhrs 2)

    try:
        plannedhrsthree = browser.find_element(By.ID, "m5e4b62f0_tdrow_[C:9]_txt-tb[R:2]").get_attribute("value");
    except:
        plannedhrsthree = 'no data'
    else:
        plannedhrs3 = plannedhrsthree[0]

    try:
        plannedhrs3 = int(plannedhrs3)
    except:
        print('no data')

    try:
        plannedhrs3 = int(plannedhrsthree[0] + plannedhrsthree[1])
    except:
        print('less than 10 hours or still no data')
        
    try:
        plannedhrs = plannedhrs + plannedhrs3
    except:
        print('no data for plannedhrs3')


        
#write to excel and move on to next wo
    sheet['A' + str(5+i)].value = WO
    sheet['B' + str(5+i)].value = jobplan
    sheet['C' + str(5+i)].value = description
    sheet['E' + str(5+i)].value = plannedhrs
    sheet['D' + str(5+i)].value = tstart
    sheet['F' + str(5+i)].value = asset
        
    nextElem = browser.find_element(By.ID, "toolactions_NEXT-tbb_image")
    type(nextElem)
    time.sleep(3)
    nextElem.click()
    wotabElem = browser.find_element(By.ID, "mbf28cd64-tab_anchor")
    type(wotabElem)
    time.sleep(3)
    wotabElem.click()
    time.sleep(5)
    plannedhrs2 = 0
    plannedhrs3 = 0
    i=i+1


