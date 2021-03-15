#Date			Programmer			Description
#-----			----------			------------
#01-09-2019		Madhu Anandam		This will generate an automated status report Monthly for the deployments.

#Importing the Libraries
import xlrd, datetime, sys, os
from datetime import timedelta, time
#from pandas import DataFrame

#Local Variables declaration

startdatelist = []
distsysCsid, distuatCsid, distoatCsid, distliveCsid, setDate = set(), set(), set(), set(), set()
listsysCsid, listDate, listDate2 = [], [], []
setsysCsid, setuatCsid, setoatCsid, setliveCsid = set(), set(), set(), set()
sysCount, uatCount, oatCount, liveCount = 0, 0, 0, 0
finalsyscountcsid, finaluatcountcsid, finaloatcountcsid, finallivecountcsid = 0, 0, 0, 0
finalsysOutcountcsid, finaluatOutcountcsid, finaloatOutcountcsid, finalliveOutcountcsid = 0, 0, 0, 0
livecsid, syscsid, uatcsid, oatcsid = set(), set(), set(), set()
syscomponent, uatcomponent, oatcomponent, livecomponent = [], [], [], []
pendSyscsid, pendUatcsid, pendOatcsid, pendLivecsid = set(), set(), set(), set()
pendSyscomponent, pendUatcomponent, pendOatcomponent, pendLivecomponent = [], [], [], []
SQLPLSQlCount, D2KCount, UnixCount, XMLCount, ADFCount, APPSCount, PortalCount, DiscovererCount, SOACount, SDFCount, MuleCount, OthersCount = [],[],[],[],[],[],[],[],[],[],[],[]

sysOutCsid, uatOutCsid, oatOutCsid, liveOutCsid = set(), set(), set(), set()

#Inputing the values

#file_loc=input('Enter the Tracker location: ')
loc = ('T:\Configuration and Release Management\Deployment Trackers\From Jan-2021\Deployment Tracker.xls')
#reportName =str(input('Press 1 for MSR; 2 for WSR; 3 for DSR: ' ))
book = xlrd.open_workbook(loc)
sheet = book.sheet_by_index(0)
lastrow = sheet.nrows
#st1 = str(input("Enter the start date in format dd/mm/yy: "))
date=(datetime.datetime.now()-timedelta(days=31)).strftime('%d/%m/%y')
start_date = datetime.datetime.strptime(date, '%d/%m/%y')
#print(start_date)
for i in range(0, 31):
        modified_date = datetime.datetime.strftime(start_date, "%d/%m/%y")
        startdatelist.append(modified_date)
        start_date = start_date + timedelta(days=1)
        #print(datetime.datetime.strftime(modified_date, "%d/%m/%y"))


'''
if reportName == '1':
    for i in range(0, 31):
        modified_date = datetime.datetime.strftime(start_date, "%d/%m/%y")
        startdatelist.append(modified_date)
        start_date = start_date + timedelta(days=1)
        #print(datetime.datetime.strftime(modified_date, "%d/%m/%y"))
elif reportName== '2':
    for i in range(0, 5):
        modified_date = datetime.datetime.strftime(start_date, "%d/%m/%y")
        startdatelist.append(modified_date)
        start_date = start_date + timedelta(days=1)
        #print(datetime.datetime.strftime(modified_date, "%d/%m/%y"))
elif reportName== '3':
    for i in range(0, 1):
        modified_date = datetime.datetime.strftime(start_date, "%d/%m/%y")
        startdatelist.append(modified_date)
        start_date = start_date + timedelta(days=1)
else:
    print('You entered a wrong report name')
    exit()

'''




sysuat=['SYS', 'UAT']
uatoat=['UAT', 'OAT']
sysoat=['SYS', 'OAT']
sysuatoat=['SYS','UAT', 'OAT']


#print(startdatelist)
for date in sorted(startdatelist):
    for rowx in range(0,lastrow ):
        datecell = sheet.cell_value(rowx, colx=0)
        try:
            datecell_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(datecell, book.datemode))
            if (str(date) == str(datecell_as_datetime.strftime('%d/%m/%y'))):
                #print(sheet.cell_value(rowx, 9))
                if 'Deployed in SYS' == sheet.cell_value(rowx, 4):
                    for csidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                        syscsid.add(csidSplit)
                        sysOutCsid.add(csidSplit) if 'OUT' == sheet.cell_value(rowx, 7) else print('',end='')
                    syscomponent.append(sheet.cell_value(rowx, 3))

                elif 'Deployed in UAT' == sheet.cell_value(rowx, 4):
                    for csidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                        uatcsid.add(csidSplit)
                        uatOutCsid.add(csidSplit) if 'OUT' == sheet.cell_value(rowx, 7) else print('',end='')
                    uatcomponent.append(sheet.cell_value(rowx, 3))

                elif 'Deployed in OAT' == sheet.cell_value(rowx, 4):
                    for csidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                        oatcsid.add(csidSplit)
                        oatOutCsid.add(csidSplit) if 'OUT' == sheet.cell_value(rowx, 7) else print('',end='')
                    oatcomponent.append(sheet.cell_value(rowx, 3))

                elif 'Deployed in LIVE' == sheet.cell_value(rowx, 4):
                    for csidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                        livecsid.add(csidSplit)
                        liveOutCsid.add(csidSplit) if 'OUT' == sheet.cell_value(rowx, 7) else print('',end='')
                    livecomponent.append(sheet.cell_value(rowx, 3))

                elif all(x in sheet.cell_value(rowx, 4) for x in sysoat):
                    for syscsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                        syscsid.add(syscsidSplit)
                        sysOutCsid.add(csidSplit) if 'OUT' == sheet.cell_value(rowx, 7) else print('',end='')
                    syscomponent.append(sheet.cell_value(rowx, 3))
                    for oatcsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                        oatcsid.add(syscsidSplit)
                        oatOutCsid.add(csidSplit) if 'OUT' == sheet.cell_value(rowx, 7) else print('',end='')
                    oatcomponent.append(sheet.cell_value(rowx, 3))

                elif all(x in sheet.cell_value(rowx, 4) for x in sysuat):
                    for syscsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                        syscsid.add(syscsidSplit)
                        sysOutCsid.add(csidSplit) if 'OUT' == sheet.cell_value(rowx, 7) else print('',end='')
                    syscomponent.append(sheet.cell_value(rowx, 3))

                    for uatcsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                        uatcsid.add(uatcsidSplit)
                        uatOutCsid.add(csidSplit) if 'OUT' == sheet.cell_value(rowx, 7) else print('',end='')
                    uatcomponent.append(sheet.cell_value(rowx, 3))

                elif all(x in sheet.cell_value(rowx, 4) for x in uatoat):
                    for uatcsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                        uatcsid.add(uatcsidSplit)
                        uatOutCsid.add(csidSplit) if 'OUT' == sheet.cell_value(rowx, 7) else print('',end='')
                    uatcomponent.append(sheet.cell_value(rowx, 3))

                    for oatcsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                        oatcsid.add(syscsidSplit)
                        oatOutCsid.add(csidSplit) if 'OUT' == sheet.cell_value(rowx, 7) else print('',end='')
                    oatcomponent.append(sheet.cell_value(rowx, 3))

                elif all(x in sheet.cell_value(rowx, 4) for x in sysuatoat):
                    for syscsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                        syscsid.add(syscsidSplit)
                        sysOutCsid.add(csidSplit) if 'OUT' == sheet.cell_value(rowx, 7) else print('',end='')
                    syscomponent.append(sheet.cell_value(rowx, 3))

                    for uatcsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                        uatcsid.add(uatcsidSplit)
                        uatOutCsid.add(csidSplit) if 'OUT' == sheet.cell_value(rowx, 7) else print('',end='')
                    uatcomponent.append(sheet.cell_value(rowx, 3))

                    for oatcsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                        oatcsid.add(syscsidSplit)
                        oatOutCsid.add(csidSplit) if 'OUT' == sheet.cell_value(rowx, 7) else print('',end='')
                    oatcomponent.append(sheet.cell_value(rowx, 3))

                elif sheet.cell_value(rowx, 4) in ['Pending in SYS', 'Error in SYS']:
                    for pendingsyscsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                        pendSyscsid.add(pendingsyscsidSplit)
                    pendSyscomponent.append(sheet.cell_value(rowx, 3))

                elif sheet.cell_value(rowx, 4) in ['Pending in UAT', 'Error in UAT']:
                    for pendinguatcsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                        pendUatcsid.add(pendinguatcsidSplit)
                    pendUatcomponent.append(sheet.cell_value(rowx, 3))

                elif sheet.cell_value(rowx, 4) in ['Pending in OAT', 'Error in OAT']:
                    for pendingoatcsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                        pendOatcsidatcsid.add(pendingoatcsidSplit)
                    pendOatcomponent.append(sheet.cell_value(rowx, 3))

                elif sheet.cell_value(rowx, 4) in ['Pending in LIVE', 'Error in LIVE']:
                    for pendinglivecsidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                        pendLivecsid.add(pendinglivecsidSplit)
                    pendLivecomponent.append(sheet.cell_value(rowx, 3))


            if (str(date) == str(datecell_as_datetime.strftime('%d/%m/%y'))):

                if sheet.cell_value(rowx, 5) in ['Type']:
                    print()

                elif sheet.cell_value(rowx, 5).lower() in [('SQL').lower(), ('PL/SQL').lower(), ('PLSQL').lower()] and 'Deployed in LIVE' in sheet.cell_value(rowx, 4):
                    SQLPLSQlCount.append(sheet.cell_value(rowx, 5))

                elif sheet.cell_value(rowx, 5).lower() in [('D2K').lower(), ('FMB').lower(), ('RDF').lower()] and 'Deployed in LIVE' in sheet.cell_value(rowx, 4):
                    D2KCount.append(sheet.cell_value(rowx, 5))

                elif sheet.cell_value(rowx, 5).lower() in [('Unix').lower()] and 'Deployed in LIVE' in sheet.cell_value(rowx, 4):
                    UnixCount.append(sheet.cell_value(rowx, 5))

                elif sheet.cell_value(rowx, 5).lower() in [('RTF').lower()] and 'Deployed in LIVE' in sheet.cell_value(rowx, 4):
                    XMLCount.append(sheet.cell_value(rowx, 5))

                elif sheet.cell_value(rowx, 5).lower() in [('ADF').lower()] and 'Deployed in LIVE' in sheet.cell_value(rowx, 4):
                    ADFCount.append(sheet.cell_value(rowx, 5))

                elif sheet.cell_value(rowx, 5).lower() in [('APPS').lower(), ('config').lower()] and 'Deployed in LIVE' in sheet.cell_value(rowx, 4):
                    APPSCount.append(sheet.cell_value(rowx, 5))

                elif sheet.cell_value(rowx, 5).lower() in [('Portal').lower()] and 'Deployed in LIVE' in sheet.cell_value(rowx, 4):
                    PortalCount.append(sheet.cell_value(rowx, 5))

                elif sheet.cell_value(rowx, 5).lower() in [('Discoverer').lower()] and 'Deployed in LIVE' in sheet.cell_value(rowx, 4):
                    DiscovererCount.append(sheet.cell_value(rowx, 5))

                elif sheet.cell_value(rowx, 5).lower() in [('SOA').lower()] and 'Deployed in LIVE' in sheet.cell_value(rowx, 4):
                    SOACount.append(sheet.cell_value(rowx, 5))

                elif sheet.cell_value(rowx, 5).lower() in [('SDF').lower()] and 'Deployed in LIVE' in sheet.cell_value(rowx, 4):
                    SDFCount.append(sheet.cell_value(rowx, 5))

                elif sheet.cell_value(rowx, 5).lower() in [('Mule').lower()] and 'Deployed in LIVE' in sheet.cell_value(rowx, 4):
                    MuleCount.append(sheet.cell_value(rowx, 5))

                elif sheet.cell_value(rowx, 4) in ['Deployed in LIVE']:
                    OthersCount.append(sheet.cell_value(rowx, 5))

                else:
                    nothing
        # startdate = startdate + datetime.timedelta(days=1)
        except:
            print('', end='')

    #window=window-1
    #print(sysOutCsid)
    #print(uatOutCsid)
    #print(oatOutCsid)
    #print(liveOutCsid)

    finalsyscountcsid = finalsyscountcsid + len(syscsid)
    finaluatcountcsid = finaluatcountcsid + len(uatcsid)
    finaloatcountcsid = finaloatcountcsid + len(oatcsid)
    finallivecountcsid = finallivecountcsid + len(livecsid)

    finalsysOutcountcsid = finalsysOutcountcsid + len(sysOutCsid)
    finaluatOutcountcsid = finaluatOutcountcsid + len(uatOutCsid)
    finaloatOutcountcsid = finaloatOutcountcsid + len(oatOutCsid)
    finalliveOutcountcsid = finalliveOutcountcsid + len(liveOutCsid)

    syscsid.clear()
    uatcsid.clear()
    oatcsid.clear()
    livecsid.clear()

    sysOutCsid.clear()
    uatOutCsid.clear()
    oatOutCsid.clear()
    liveOutCsid.clear()
    #print(finalsysOutcountcsid)

finalsyscountcsidList, finaluatcountcsidList, finaloatcountcsidList, finallivecountcsidList=[],[],[],[]
finalsyscountcsidList.append(finalsyscountcsid)
finaluatcountcsidList.append(finaluatcountcsid)
finaloatcountcsidList.append(finaloatcountcsid)
finallivecountcsidList.append(finallivecountcsid)

os.chdir("T:\\Configuration and Release Management\\Status Reports\\MSR\\Component_Reports") #Changing the current directory
sys.stdout=open("MSR_Deployment_Tracker_Report_"+str((datetime.datetime.now()).strftime('%d-%m-%y'))+".txt","w") #Creating a file to store the output data

print("Deployed Changes as CSID and Component wise", end='\n\n')
print("SYS "'\t\t'"UAT"'\t\t'"OAT"'\t\t'"LIVE")
print(finalsyscountcsid, '\t\t', finaluatcountcsid, '\t\t', finaloatcountcsid, '\t\t', finallivecountcsid)
print(len(syscomponent), '\t\t', len(uatcomponent), '\t\t', len(oatcomponent), '\t\t', len(livecomponent))
print("", end='\n\n')
print("--------------------------------------------------------")
print("Pending Components", end='\n\n')
print("SYS "'\t\t'"UAT"'\t\t'"OAT"'\t\t'"LIVE")
print(len(pendSyscomponent), '\t\t', len(pendUatcomponent), '\t\t', len(pendOatcomponent), '\t\t',
      len(pendLivecomponent))

print("", end='\n\n')
print("--------------------------------------------------------")
print("out of window deployment CSID count", end='\n\n')
print("SYS "'\t\t'"UAT"'\t\t'"OAT"'\t\t'"LIVE")
print(finalsysOutcountcsid, '\t\t', finaloatOutcountcsid, '\t\t', finaloatOutcountcsid, '\t\t', finalliveOutcountcsid)
print("", end='\n\n')
print("--------------------------------------------------------")
print("Technology Components deployed in LIVE", end='\n\n')
print('SQL/PLSQL Count is       :',len(SQLPLSQlCount))
print('D2K Count is             :',len(D2KCount))
print('Unix Count is            :',len(UnixCount))
print('XML Count is             :',len(XMLCount))
print('ADF Count is             :',len(ADFCount))
print('Oracle APPS Count is     :',len(APPSCount))
print('Portal Count is          :',len(PortalCount))
print('Discoverer Count is      :',len(DiscovererCount))
print('SOA Count is             :',len(SOACount))
print('SDF Count is             :',len(SDFCount))
print('Mule Count is            :',len(MuleCount))
print('Others Count is          :',len(OthersCount))
print('Total count is           :', (len(SQLPLSQlCount)+len(D2KCount)+len(UnixCount)+len(XMLCount)+len(ADFCount)+len(APPSCount)+
                                     len(PortalCount)+len(DiscovererCount)+len(SOACount)+len(SDFCount)+len(MuleCount)+len(OthersCount)))

#dataframe = DataFrame({'SYS':[finalsyscountcsidList,len(syscomponent)], 'UAT':[finaluatcountcsidList,len(uatcomponent)],'OAT':[finaloatcountcsidList,len(oatcomponent)], 'LIVE':[finallivecountcsidList,len(livecomponent)]})
#dataframe = DataFrame([final_list_service])
#print(dataframe)
#print(len(final_list_service))
#print(final_list_service[1][1])
#dataframe.to_excel('test.xlsx', sheet_name='sheet1', index=False)

sys.stdout.close()