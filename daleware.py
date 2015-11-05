#!/user/bin/python3
# -*- coding: utf-8 -*-
# compatible with Python 3.4.3

__author__="Bo Zhou"
__copyright__ = "Copyright 2015, Dalewareriverkeeper project "
__credits__ = ["Bo Zhou"]
__license__ = "MIT"
__version__ = "1.0.0"
__maintainer__ = "Bo Zhou"
__email__ = "bzhou2@ualberta.ca"
__status__ = "Testing"

import time
import bs4
import xlsxwriter
import requests
    
def show_time(time):
    hours = time//3600
    minutes = (time//60)%60
    seconds = time%60
    print ("program runs for "+str(int(hours))+" hours, "+str(int(minutes))+" minutes, "+str(seconds)+" seconds.")

def write_to_excel(allData):
    wb = xlsxwriter.Workbook("test.xlsx")
    ws = wb.add_worksheet('Event Data') 
    ws.set_column("A:G",40)
    ws.set_column("G:G",100)
    cellFormat = wb.add_format({'text_wrap': True})
    ws
    row = 0
    for data in allData:
        column = 0
        for el in data:
            ws.write(row, column, el, cellFormat)
            column += 1            
        row += 1
    wb.close()
    return

# get whole page data
# return: [startDate, EndDate, Url, name, time, location, description]
def catch_data(response):
    r = response.url
    soup = bs4.BeautifulSoup(response.content)
    pageContent = soup.find(id = "pageContent")
    allInfo = pageContent.find_all("tr")
    result = []
    # deal with date:
    dateList = allInfo[1].text.split()
    result.append(dateList[1])
    try:
        result.append(dateList[3])
    except IndexError:
        result.append(dateList[1])
    # add event url
    result.append(r)
    # add event name
    result.append(allInfo[0].text)
    # add event time
    timeList = allInfo[2].text.split()
    timeList.pop(0)
    time = " ".join(timeList)
    result.append(time)
    # add location:
    locationList = allInfo[3].text.split()
    for i in range(2):
        locationList.pop(0)
    location = " ".join(locationList)
    result.append(location)
    # add description:
    result.append(allInfo[4].text)
    return result

if __name__ == '__main__':
    startTime = time.time()
    saveFileName = "Daleware"
    crawlUrl = 'http://www.delawareriverkeeper.org/about/event.aspx?Id='
    pageId = 0
    allData = []
    nodata = 0
    while True:
        pageUrl = crawlUrl + str(pageId)
        try:
            response = requests.get(pageUrl)
        except Exception as e:
            print("cannot open")
        if response.status_code == 200:
            pageData = catch_data(response)
            allData.append(pageData)
            print ("opening page: "+str(pageId))
            nodata = 0
        elif response.status_code == 500:
            print ("No result on page: "+str(pageId))
            nodata += 1
        else:
            print ("Cannot open page: "+str(pageId))
        pageId += 1
        if nodata == 26:
            break
    write_to_excel(allData)
    elapsedTime = time.time() - startTime
    show_time(elapsedTime)   
    print('all finish! ')