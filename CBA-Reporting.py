#Altered line 153 in runEWlaunch() to not wait for ewlauch program exit before terminating script
import serial
import io
import re
import datetime
import time
import os
import subprocess
import serial.tools.list_ports

#Delete the report if it's already been run today.     
def initReportFolder():
    print 'Initializing the reporting folder.....',
    if os.path.exists(RAW_REPORT_OUT+filename):
        os.remove(RAW_REPORT_OUT+filename)
    print 'DONE'
    
#Find the CBA COM port number via description 'USB Serial Port' (and NOT 'Jamex USB Serial Port')
def findComPort():
    print 'Finding CBA connection................',
    usbser = 'USB Serial Port'
    jam = 'Jamex'
    ports = list(serial.tools.list_ports.comports())
    for p in ports:
        for i in range(0,len(p)):
            if (p[i].find(usbser)) > -1 and (p[i].find(jam)) == -1:
                COM_PORT = p[i-1]
    print 'DONE'
    return COM_PORT

#Connect via Serial USB to the CBA
def serialConnect(COM_PORT, filename):    
    ser = serial.Serial()
    ser.port = COM_PORT
    ser.baudrate = 1200
    ser.parity = serial.PARITY_NONE
    ser.bytesize = 8
    ser.stopbits = 1
    ser.rtscts = True
    ser.timeout = None
    ser.open()

    if ser.isOpen()==True:
        print ('Connected to CBA on ' + COM_PORT)
    f=open(RAW_REPORT_OUT + filename,'w')
    loop=True
    print ('**Please press the YES and NO buttons on the CBA circuitboard to start**')
    while loop==True:
        line = ser.readline()
        print line
        print>>f,line
        if line.find('End of Report') > -1:
            time.sleep(1)
            loop=False
    print>>f,today
    print>>f,ctime
    
    f.close()
    ser.close()

def parseReport(filepath):
    print 'Parsing ' + os.path.basename(filepath)
    rawList=[]
    outList=[]
    f=open(filepath,'r')
    for line in f:
        temp = re.findall(r'([A-z0-9/.:]+)',line)     
        rawList.append(temp)
    for el in rawList:
        if el:
            if 'PuTTY' in el:
		continue
	    else:
                outList.append(el)
    ##Searching for the time in the filepath name could be problematic - I removed this?******************************************************************* 
    date=os.path.basename(filepath)[:10]
    time=os.path.basename(filepath)[11:17]
    outList.append(date)
    outList.append(time)
    return outList

#Find the newest file in a particular directory. Used to import previous report data before importing current report. Uses date created, as opposed to date modified.
def findNewestFile(path):
    files = []
    for f in os.listdir(path):
        if os.path.isfile(path+f):
            files.append(path+f)
    newest = max(files , key = os.path.getctime)
    return newest

#Create a dictionary of the report data
def createDict(reportData,dictName):
    dictName['SN'] = reportData[0][5]
    dictName['Copy-Cash-A4BW'] = int(reportData[7][2])
    dictName['Copy-Cash-A3BW'] = int(reportData[8][2])
    dictName['Copy-Cash-A4Col'] = int(reportData[9][2])
    dictName['Copy-Cash-A3Col'] = int(reportData[10][2])
    dictName['Copy-Byp-A4BW'] = int(reportData[11][2])
    dictName['Copy-Byp-A3BW'] = int(reportData[12][2])
    dictName['Copy-Byp-A4Col'] = int(reportData[13][2])
    dictName['Copy-Byp-A3Col'] = int(reportData[14][2])
    dictName['Print-Byp'] = float(reportData[16][0])
    dictName['Print-Cash'] = float(reportData[18][0])
    dictName['$50-Cnt'] = int(reportData[26][1])
    dictName['$20-Cnt'] = int(reportData[28][1])
    dictName['$10-Cnt'] = int(reportData[30][1])
    dictName['$5-Cnt'] = int(reportData[32][1])
    dictName['Change-Tube'] = float(reportData[35][0])
    dictName['Coin-Drawer'] = float(reportData[22][0])
    dictName['Copy-Price-A4BW'] = float(reportData[36][4])
    dictName['Copy-Price-A3BW'] = float(reportData[37][2])
    dictName['Copy-Price-A4Col'] = float(reportData[38][2])
    dictName['Copy-Price-A3Col'] = float(reportData[39][2])
    dictName['Date'] = reportData[41]
    dictName['Time'] = reportData[42]

#Calculate totals / variances needed for report output
def calcReportFields(rOld,rNew):
    print 'Calculating Report Fields'
    reportVar={}
    reportVar['CCash-A4BW-Var'] = rNew['Copy-Cash-A4BW'] - rOld['Copy-Cash-A4BW']
    reportVar['CCash-A3BW-Var'] = rNew['Copy-Cash-A3BW'] - rOld['Copy-Cash-A3BW']
    reportVar['CCash-A4Col-Var'] = rNew['Copy-Cash-A4Col'] - rOld['Copy-Cash-A4Col']
    reportVar['CCash-A3Col-Var'] = rNew['Copy-Cash-A3Col'] - rOld['Copy-Cash-A3Col']
    reportVar['Copy-A4BW-Tot'] = reportVar['CCash-A4BW-Var'] * rNew['Copy-Price-A4BW']
    reportVar['Copy-A3BW-Tot'] = reportVar['CCash-A3BW-Var'] * rNew['Copy-Price-A3BW']
    reportVar['Copy-A4Col-Tot'] = reportVar['CCash-A4Col-Var'] * rNew['Copy-Price-A4Col']
    reportVar['Copy-A3Col-Tot'] = reportVar['CCash-A3Col-Var'] * rNew['Copy-Price-A3Col']
    reportVar['CCash-Tot'] = reportVar['Copy-A4BW-Tot'] + reportVar['Copy-A3BW-Tot'] + reportVar['Copy-A4Col-Tot'] + reportVar['Copy-A3Col-Tot']
    reportVar['BCash-A4BW-Var'] = rNew['Copy-Byp-A4BW'] - rOld['Copy-Byp-A4BW']
    reportVar['BCash-A3BW-Var'] = rNew['Copy-Byp-A3BW'] - rOld['Copy-Byp-A3BW']
    reportVar['BCash-A4Col-Var'] = rNew['Copy-Byp-A4Col'] - rOld['Copy-Byp-A4Col']
    reportVar['BCash-A3Col-Var'] = rNew['Copy-Byp-A3Col'] - rOld['Copy-Byp-A3Col']
    reportVar['PCash-Tot'] = rNew['Print-Cash'] - rOld['Print-Cash']
    reportVar['PByp-Tot'] = rNew['Print-Byp'] - rOld['Print-Byp']
    reportVar['$50-Tot'] = rNew['$50-Cnt'] * 50
    reportVar['$20-Tot'] = rNew['$20-Cnt'] * 20
    reportVar['$10-Tot'] = rNew['$10-Cnt'] * 10
    reportVar['$5-Tot'] = rNew['$5-Cnt'] * 5
    reportVar['Coin-Drawer'] = rNew['Coin-Drawer']
    reportVar['Coin-Tube'] = rNew['Change-Tube']
    reportVar['Expected-Cash'] = reportVar['CCash-Tot'] + reportVar['PCash-Tot'] + FLOAT
    return reportVar

#Print the report to the default printer
def printTxtFile(filepath):
    print 'Printing report'
    subprocess.call(["notepad","/p",filepath], shell=False)

def runEWlaunch():
    ewPath = "C:/Program Files/EnvisionWare/ewlaunch/ewlaunch.exe"
    # p= subprocess.Popen(ewPath, shell=True, stdout = subprocess.PIPE)
    p = subprocess.Popen(ewPath, shell=True, stdin=None, stdout=None, stderr=None)

##BEGIN
print '************************************************'
print '*     CBA-V Serial Port Reporting Script       *'
print '************************************************'


COM_PORT = findComPort()

FLOAT = 100
today = datetime.datetime.today().strftime('%d-%m-%Y')
ctime = datetime.datetime.now().time().strftime('%H%M%S')
filename = today+'_'+COM_PORT+'.txt'
RAW_REPORT_OUT = 'C:/Reports/'

initReportFolder()

oldRepFile = findNewestFile(RAW_REPORT_OUT)
print 'Previous report file : ' + os.path.basename(oldRepFile)
oldReport = parseReport(oldRepFile)
oldDict = {}
createDict(oldReport,oldDict)

serialConnect(COM_PORT, filename)

newRepFile = RAW_REPORT_OUT + filename
newReport = parseReport(newRepFile)
newDict={}
createDict(newReport,newDict)

v = {}
v = calcReportFields(oldDict,newDict)

    
##Generate Report
today = datetime.datetime.today().strftime('%a %d %b %Y')
ctime = datetime.datetime.now().time().strftime('%H:%M')

reportFile = 'C:/PrintedReports/'+today+'.txt'

f=open(reportFile,'w')
print>>f,today+' - '+ctime
print>>f,'Envisionware	CBA-V	S/N:'+newDict['SN']
print>>f,'--------------------------------------------------------------------------------'
print>>f,'Copy Meter Counts'
print>>f,'--------------------------------------------------------------------------------'
print>>f,'Cash'
print>>f,'\tPREV\tCURR\tVAR\tCOST\tTOTAL'
print>>f,'A4-BW\t'+str(oldDict['Copy-Cash-A4BW'])+'\t'+str(newDict['Copy-Cash-A4BW'])+'\t'+str(v['CCash-A4BW-Var'])+'\t$'+str("{0:.2f}".format(newDict['Copy-Price-A4BW']))+'\t$'+str("{:.2f}".format(v['Copy-A4BW-Tot']))
print>>f,'A3-BW\t'+str(oldDict['Copy-Cash-A3BW'])+'\t'+str(newDict['Copy-Cash-A3BW'])+'\t'+str(v['CCash-A3BW-Var'])+'\t$'+str("{0:.2f}".format(newDict['Copy-Price-A3BW']))+'\t$'+str("{:.2f}".format(v['Copy-A3BW-Tot']))
print>>f,'A4-Col\t'+str(oldDict['Copy-Cash-A4Col'])+'\t'+str(newDict['Copy-Cash-A4Col'])+'\t'+str(v['CCash-A4Col-Var'])+'\t$'+str("{0:.2f}".format(newDict['Copy-Price-A4Col']))+'\t$'+str("{:.2f}".format(v['Copy-A4Col-Tot']))
print>>f,'A3-Col\t'+str(oldDict['Copy-Cash-A3Col'])+'\t'+str(newDict['Copy-Cash-A3Col'])+'\t'+str(v['CCash-A3Col-Var'])+'\t$'+str("{0:.2f}".format(newDict['Copy-Price-A3Col']))+'\t$'+str("{:.2f}".format(v['Copy-A3Col-Tot']))
print>>f,'\nBypass'
print>>f,'\tPREV\tCURR\tVAR'
print>>f,'A4-BW\t'+str(oldDict['Copy-Byp-A4BW'])+'\t'+str(newDict['Copy-Byp-A4BW'])+'\t'+str(v['BCash-A4BW-Var'])
print>>f,'A3-BW\t'+str(oldDict['Copy-Byp-A3BW'])+'\t'+str(newDict['Copy-Byp-A3BW'])+'\t'+str(v['BCash-A3BW-Var'])
print>>f,'A4-Col\t'+str(oldDict['Copy-Byp-A4Col'])+'\t'+str(newDict['Copy-Byp-A4Col'])+'\t'+str(v['BCash-A4Col-Var'])
print>>f,'A3-Col\t'+str(oldDict['Copy-Byp-A3Col'])+'\t'+str(newDict['Copy-Byp-A3Col'])+'\t'+str(v['BCash-A3Col-Var'])
print>>f,'--------------------------------------------------------------------------------'
print>>f,'COPY CASH TOTAL\t\t\t\t$'+str("{:.2f}".format(v['CCash-Tot']))
print>>f,'================================================================================\n'
print>>f,'Printing Charges'
print>>f,'--------------------------------------------------------------------------------'
print>>f,'\tPREV\tCURR\tVAR'
print>>f,'Bypass\t$'+str("{:.2f}".format(oldDict['Print-Byp']))+'\t$'+str("{:.2f}".format(newDict['Print-Byp']))+'\t$'+str("{:.2f}".format(v['PByp-Tot']))
print>>f,'Cash\t$'+str("{:.2f}".format(oldDict['Print-Cash']))+'\t$'+str("{:.2f}".format(newDict['Print-Cash']))+'\t$'+str("{:.2f}".format(v['PCash-Tot']))
print>>f,'--------------------------------------------------------------------------------'
print>>f,'PRINTING CASH TOTAL\t\t\t$'+str("{:.2f}".format(v['PCash-Tot']))
print>>f,'================================================================================\n'
print>>f,'--------------------------------------------------------------------------------'
print>>f,'Note Counts\tCNT\tTOTAL'
print>>f,'--------------------------------------------------------------------------------'
print>>f,'$20\t\t'+str(newDict['$20-Cnt'])+'\t$'+str("{:.2f}".format(v['$20-Tot']))
print>>f,'$10\t\t'+str(newDict['$10-Cnt'])+'\t$'+str("{:.2f}".format(v['$10-Tot']))
print>>f,'$5\t\t'+str(newDict['$5-Cnt'])+'\t$'+str("{:.2f}".format(v['$5-Tot']))
print>>f,'================================================================================\n'
print>>f,'--------------------------------------------------------------------------------'
print>>f,'CBA Meter Readings'
print>>f,'--------------------------------------------------------------------------------'
print>>f,'CoinDrawer\t\t$'+str("{:.2f}".format(v['Coin-Drawer']))
print>>f,'CoinTubes\t\t$'+str("{:.2f}".format(v['Coin-Tube']))
print>>f,'================================================================================\n'
print>>f,'EXPECTED CASH = COPY CASH TOTAL($'+str("{:.2f}".format(v['CCash-Tot']))+') + PRINT CASH TOTAL($'+str("{:.2f}".format(v['PCash-Tot']))+') + FLOAT($'+str("{:.2f}".format(FLOAT))+')'
print>>f,'EXPECTED CASH = $'+str("{:.2f}".format(v['CCash-Tot']))+' + $'+str("{:.2f}".format(v['PCash-Tot']))+' + $'+str("{:.2f}".format(FLOAT))+' = $'+ str("{:.2f}".format(v['Expected-Cash']))
print>>f,'================================================================================\n'
print>>f,'I have verified that all cash has been bagged and float levels are correct.'
print>>f,'Coin Machine ID: (STK-WP/STK-IT/AP/EH/EHHC/MP/PM)\n'
print>>f,'Tag No.:\n'
print>>f,'Staff Name and Signature:\n\n\n\n'
print>>f,'Guard Name and Signature:'
f.close()

printTxtFile(reportFile)
runEWlaunch()

