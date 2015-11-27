__author__ = 'tzonliang.hsieh'
from xlwt import easyxf
from xlutils.copy import copy
from xlrd import open_workbook
import time
import os

class Excelxlrw:
    def __init__(self):
        print "Init class"

    #def

    def CheckFileVail(self,filepath):
        try:
            Pathindex=filepath.strip('\n')
            print ('Pathindex:',Pathindex)
            os.path.dirname(Pathindex)
            rb = open_workbook(os.path.abspath(filepath), formatting_info=True)
            return True
        except Exception as inst:
            print "open fail on:"+inst.strerror
            return False

    def ModifyExcelData(self,filepath,sheet,col,row,data):
        try:

            #open file and modify
            filepath=filepath.strip('\n')
            #print ('Pathindex:',Pathindex)
            os.path.dirname(filepath)
            ProjectSVNPath = os.path.abspath(filepath)
            rb = open_workbook(ProjectSVNPath, formatting_info=True)
            rs = rb.sheet_by_index(sheet)
            wb = copy(rb)
            ws = wb.get_sheet(sheet)
            ws.write(col,row,data)
            os.remove(ProjectSVNPath)
            wb.save(ProjectSVNPath)
            return True
        except Exception as inst:
            print "Fail on :"+inst.strerror


Times = time.time()
CheckDateStr = 'Check Date:' + time.strftime('%Y/%m/%d',time.localtime(Times))
OpenExcelObj = Excelxlrw()
#print OpenExcelObj.__class__
if OpenExcelObj.CheckFileVail('scorecard.xls'):
    OpenExcelObj.ModifyExcelData('scorecard.xls',sheet=0,col=1,row=3,data=CheckDateStr)


#wb = Workbook('ScoreCard.xls')
#rb = open_workbook('ScoreCard.xls', formatting_info=True)
#rs = rb.sheet_by_index(0)
#wb = copy(rb)
#ws = wb.get_sheet(0)

#ws.write(1,3,CheckDateStr)

#os.remove('ScoreCard.xls')
#cell = sheet.cell(1,3)
#sheet.write(1,3,CheckDateStr)
#wb.save('ScoreCard.xls')
#print cell
#print cell.value


