import webbrowser
import time
from datetime import datetime
import xlrd

progPath = ("SchoolProgram.xlsx")
wb = xlrd.open_workbook(progPath)
sheet = wb.sheet_by_index(0)
hour = 0
timeBefore = int(sheet.cell_value(21, 2))
Lessons= ["French", "PE", "ComSci", "AncGr", "Geom", "Eng", "Alg", "Relig", "His", "Chem", "SocSt", "Lit", "Phys", "Bio", "Lang"]

def readHour(row, col):
    start, end = sheet.cell_value(row, col).split("-")
    startH, startM, startS = start.split(":")
    return [int(startH), int(startM), int(startS)]

schTime1 = readHour(1, 0)
schTime2 = readHour(2, 0)
schTime3 = readHour(3, 0)
schTime4 = readHour(4, 0)
schTime5 = readHour(5, 0)
schTime6 = readHour(6, 0)
schTime7 = readHour(7, 0)
def defHour(time):
    if timeEqual(time, schTime1[0], schTime1[1] - timeBefore, schTime1[2]):       #8:00
        hour = 1
    elif timeEqual(time, schTime2[0], schTime2[1] - timeBefore, schTime2[2]):      #8:50
        hour = 2
    elif timeEqual(time, schTime3[0], schTime3[1] - timeBefore, schTime3[2]):     #9:40
        hour = 3
    elif timeEqual(time, schTime4[0], schTime4[1] - timeBefore, schTime4[2]):    #10:30
        hour = 4
    elif timeEqual(time, schTime5[0], schTime5[1] - timeBefore, schTime5[2]):    #11:20
        hour = 5
    elif timeEqual(time, schTime6[0], schTime6[1] - timeBefore, schTime6[2]):    #12:10
        hour = 6
    elif timeEqual(time, schTime7[0], schTime7[1] - timeBefore, schTime7[2]):    #13:00
        hour = 7
    else:
        hour = None
    if hour != None:
        print(hour)
    return hour

def timeEqual(timestr, hours, minutes, seconds):
    if minutes < 0:
        hours -= 1
        minutes = 60 + minutes
    if seconds < 0:
        minutes -= 1
        seconds = 60 + seconds
    if int(now.strftime("%H")) == hours and int(now.strftime("%M")) == minutes and int(now.strftime("%S")) == seconds:
        return True

now = datetime.now()
timestr = now.strftime("%H:%M:%S")
while int(now.strftime("%H")) < 14:
    now = datetime.now()
    timestr = now.strftime("%H:%M:%S")
    day = datetime.today().weekday() + 1
    hour = defHour(timestr)
    if hour != None:
        print("Detecting Current Lesson:")
        lesson = sheet.cell_value(hour, day)
        print(lesson)
        for x in Lessons:
            print("Searching in Lessons:")
            print(x)
            if x == lesson:
                timeBefore = int(sheet.cell_value(21, 2))
                print("Trying to open Meeting...")
                webbrowser.open(sheet.cell_value(Lessons.index(x)+1, 8))
        time.sleep(1)
