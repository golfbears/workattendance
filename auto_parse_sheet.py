import openpyxl
from openpyxl.styles import Font,Color,Fill
from openpyxl.styles.colors import RED
import time
import datetime
import logging
import re
import os


def parseTableName(abnormalCase):
    currentDir=os.getcwd()
    sheetList=os.listdir(currentDir)
    mark = 0
    for peformanceSheet in sheetList:
        sheetName=peformanceSheet.split('.')
        print(peformanceSheet)
        if sheetName[1] == 'xlsx':
            #print(sheetName[0]+'是表格')
            tableOpen=openpyxl.load_workbook(str(currentDir+'\\'+peformanceSheet),data_only=True)
            allSheetNames=tableOpen.get_sheet_names()
            print(allSheetNames)
            
            for curSheetName in allSheetNames:
                if abnormalCase in curSheetName:
                    Sheet=tableOpen.get_sheet_by_name(curSheetName)
                    print(curSheetName)
                    mark+=1
                    for row_line in range(2,6):     #skip the first line
                        lineDate=Sheet.cell(row= row_line, column=4).value
                        yearMonthDay=lineDate.split('-')
                        weekday = datetime.date(int(yearMonthDay[0]), int(yearMonthDay[1]), int(yearMonthDay[2])).weekday()
                        print(yearMonthDay, weekday)
                #if '外勤' in curSheet:
                #    print('外勤sheet名字匹配成功')
                #    Sheet=tableOpen.get_sheet_by_name(curSheet)  
                #elif '请假' in curSheet:
                #    print('请假sheet名字匹配成功')
                #elif '加班' in curSheet:
                #    print('加班sheet名字匹配成功')
                #elif '出差' in curSheet:
                #    print('出差sheet名字匹配成功')    
        else:
            print('$$$$$$$$$$$$$$$$$$$$$$$$$'+ sheetName[0]+'不是表格'+'$$$$$$$$$$$$$$$$$$$$$$$$$')
    print(type(Sheet))        
    if mark == 1:
        return Sheet
    elif mark > 1:
        return 'More'
    else:
        return 'Null'           

kaoqinji=parseTableName('考勤机')
print(kaoqinji)
#qinqi=parseTableName('外外勤')                  
