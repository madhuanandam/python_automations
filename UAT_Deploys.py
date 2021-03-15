import xlrd, datetime, sys, os
from datetime import timedelta, time



livecsid,uatcsid, uatOutCsid, liveOutCsid=set(), set(), set(), set()
livecomponent, uatcomponent=[], []
startdatelist = []



loc = ('T:\Configuration and Release Management\Deployment Trackers\From Jan-2021\Deployment Tracker.xls')

book = xlrd.open_workbook(loc)
sheet = book.sheet_by_index(0)
lastrow = sheet.nrows
date=(datetime.datetime.now()-timedelta(days=7)).strftime('%d/%m/%y')
start_date = datetime.datetime.strptime(date, '%d/%m/%y')

for i in range(0, 30):
        modified_date = datetime.datetime.strftime(start_date, "%d/%m/%y")
        startdatelist.append(modified_date)
        start_date = start_date + timedelta(days=1)


for date in sorted(startdatelist, key=lambda x: datetime.datetime.strptime(x, '%d/%m/%y')):
    for rowx in range(0,lastrow ):
        datecell = sheet.cell_value(rowx, colx=0)
        try:
            datecell_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(datecell, book.datemode))
            if (str(date) == str(datecell_as_datetime.strftime('%d/%m/%y'))):
                #print(sheet.cell_value(rowx, 9))
                if 'Deployed in LIVE' == sheet.cell_value(rowx, 4):
                    for csidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                        livecsid.add(csidSplit)
                        liveOutCsid.add(csidSplit) if 'OUT' == sheet.cell_value(rowx, 7) else print('',end='')
                    livecomponent.append(sheet.cell_value(rowx, 3))

                elif 'Deployed in UAT' == sheet.cell_value(rowx, 4):
                    for csidSplit in str(sheet.cell_value(rowx, 2)).split('/'):
                        uatcsid.add(csidSplit)
                        uatOutCsid.add(csidSplit) if 'OUT' == sheet.cell_value(rowx, 7) else print('',end='')
                    uatcomponent.append(sheet.cell_value(rowx, 3))

                else:
                    nothing
        except:
            print('', end='')


os.chdir("T:\\Configuration and Release Management\\Status Reports\\UAT\\Weekly_UAT_Deployment_Reports") #Changing the current directory
sys.stdout=open("UAT_Deployment_Status_Report_"+str((datetime.datetime.now()).strftime('%d-%m-%y'))+".txt","w") #Creating a file to store the output data

#print('-----------Components deployed in LIVE--------------------')

#for livecs in livecsid:
#    print(livecs)
#for livecomp in livecomponent:
#    print(livecomp)

#print('',end='\n')
#print('',end='\n')
#print('',end='\n')

#print('Components deployed to UAT')

#for uatcs in uatcsid:
#    print(uatcs)
#print(uatOutCsid)
#for uatcomp in uatcomponent:
#    print(uatcomp)

print('-----------------------------------------Component compare----------------------------------------------')
for livecomp in livecomponent:
    if livecomp in uatcomponent:
        print('Components deployed in LIVE & UAT                 ---> ', livecomp)

    else:
        print('Components deployed in LIVE & not deployed in UAT ---> ', livecomp)






