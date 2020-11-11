
import xlrd
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
sns.set()


#read all the excels
AZNM = xlrd.open_workbook('/projects/b1045/EVTool/AZNM.xlsx')
CAMX = xlrd.open_workbook('/projects/b1045/EVTool/CAMX.xlsx')
ERCT = xlrd.open_workbook('/projects/b1045/EVTool/ERCT.xlsx')
FRCC = xlrd.open_workbook('/projects/b1045/EVTool/FRCC.xlsx')
MROE = xlrd.open_workbook('/projects/b1045/EVTool/MROE.xlsx')
MROW = xlrd.open_workbook('/projects/b1045/EVTool/MROW.xlsx')
NEWE = xlrd.open_workbook('/projects/b1045/EVTool/NEWE.xlsx')
NWPP = xlrd.open_workbook('/projects/b1045/EVTool/NWPP.xlsx')
NYCW = xlrd.open_workbook('/projects/b1045/EVTool/NYCW.xlsx')
NYLI = xlrd.open_workbook('/projects/b1045/EVTool/NYLI.xlsx')
SRVC = xlrd.open_workbook('/projects/b1045/EVTool/SRVC.xlsx')
NYUP = xlrd.open_workbook('/projects/b1045/EVTool/NYUP.xlsx')
RFCE = xlrd.open_workbook('/projects/b1045/EVTool/RFCE.xlsx')
RFCM = xlrd.open_workbook('/projects/b1045/EVTool/RFCM.xlsx')
RFCW = xlrd.open_workbook('/projects/b1045/EVTool/RFCW.xlsx')
RMPA = xlrd.open_workbook('/projects/b1045/EVTool/RMPA.xlsx')
SPNO = xlrd.open_workbook('/projects/b1045/EVTool/SPNO.xlsx')
SPSO = xlrd.open_workbook('/projects/b1045/EVTool/SPSO.xlsx')
SRMV = xlrd.open_workbook('/projects/b1045/EVTool/SRMV.xlsx')
SRMW = xlrd.open_workbook('/projects/b1045/EVTool/SRMW.xlsx')
SRSO = xlrd.open_workbook('/projects/b1045/EVTool/SRSO.xlsx')
SRTV = xlrd.open_workbook('/projects/b1045/EVTool/SRTV.xlsx')

#create root window
#root = tk.Tk()
#root.title('eGRID')
v=0

#set inputs to variables to be called for later
enter=input('what is your zip code? ')
startyear = input('what year do you plan on starting to own the vehicle? ')
endyear = input('when do you plan to stop owning? ')
know = input('do you know how many miles you drive per year? (yes or no)')

#if they say know to knowing the miles per year, it will proceed with average miles b/c v value change will be called upon later
if know=='yes':
    global theirnumm
    theirnumm=input('how many miles do you drive per year?')
else:
    v=1


#first entry zip code
#enter = tk.Entry(root)
#enter.pack()
#enter.focus_set()
#enter.grid(row=0, column = 0)

#start year enter
#set to global so it can be called in future functions
global startEnter
#startEnter = tk.Entry(root)
#startEnter.grid(row=1, column=0)
#startEnter.focus_set()

#end year enter
#set to global so it can be called in future functions
global endEnter
#endEnter = tk.Entry(root)
#endEnter.grid(row=2, column = 0)
#endEnter.focus_set()

#create ints to use later




#function to get eGRID from zipcode
def getgrid():
    #create empty list which will get all the percents later in the function
    vList=[]
    zipGrids = xlrd.open_workbook('/projects/b1045/EVTool/zipGrid.xlsx')
    #set variable sheet as the 5th sheet in the zipGrid file
    sheet = zipGrids.sheet_by_index(4)

    for row_num in range(sheet.nrows):#sort thru all rows in excel, creating loop
        row_value = sheet.row_values(row_num)
        if row_value[1] == int(enter):#loop cycles until finds a row value where zip code is equal to the zip code in the 2nd column
            #set eGRID to global variable to be able to call upon later
            global eGRID
            #set eGRID value to the 4th column of the same row which was looped into with the code above
            eGRID = row_value[3]

    energyCalc = xlrd.open_workbook('/projects/b1045/EVTool/energyCalculations.xlsx')
    #set variable for the exact sheet in the energyCalculations excel file
    carbonDioxCalc = energyCalc.sheet_by_index(0)
    global COTotal


    #loops through new excel rows
    for row_num2 in range(carbonDioxCalc.nrows):
        row_value2=carbonDioxCalc.row_values(row_num2)
        if row_value2[1]==eGRID:#when a value in the 2nd column equals eGRID, it takes the row
            global COtotal
            COtotal=row_value2[4]#COTotal for the eGRID is the 5th column of the row defined above
            global kwhtotal
            kwhtotal=row_value2[5]#kwh total for the eGRID is the 6th column of the row defined above


    timeSheet = xlrd.open_workbook('/projects/b1045/EVTool/'+eGRID+'.xlsx')
    sheetTwo = timeSheet.sheet_by_index(0)
    #theirStart = int(startEnter.get())
    #theirEnd = int(endEnter.get())
    theirStart = 2018
    theirEnd = 2051

    while True:
        for startRow in range(sheetTwo.nrows):
            startDateRow=sheetTwo.row_values(startRow)
            if startDateRow[1]==theirStart:
                yearFuture = sheetTwo.row_values(rowx=-theirStart + 2052)
                COpercentage =((startDateRow[2] * 95.6191) + (startDateRow[4] * 53.06) + (startDateRow[3] * 74.57193) + (.1126 * startDateRow[5] * 52.04))/((yearFuture[2] * 95.6191) + (yearFuture[4] * 53.06) + (yearFuture[3] * 74.57193) + (.1126 * yearFuture[5] * 52.04))
                kwhpercentage = startDateRow[6]/yearFuture[6]
                COtotal = COtotal * COpercentage
                kwhtotal = kwhtotal * kwhpercentage
                factor = COtotal/kwhtotal
                if v==1:
                    vList.append(factor*18724.27*0.1988*2.205)
                else:
                    vList.append(factor*int(theirnumm)*.1988*2.205*1.60934)


        theirStart+=1
        if theirStart>=theirEnd:
            break

    return vList


def getCoal():
    #create empty list to append percent changes to
    coalList=[]

    zipGrids = xlrd.open_workbook('/projects/b1045/EVTool/zipGrid.xlsx')
    sheet = zipGrids.sheet_by_index(4)

    for row_num in range(sheet.nrows):#sort thru all rows in excel
        row_value = sheet.row_values(row_num)#sorts thru rows
        if row_value[1] == int(enter):#cycles until input equals a row
            global eGRID
            eGRID = row_value[3]

    energyCalc = xlrd.open_workbook('/projects/b1045/EVTool/energyCalculations.xlsx')
    carbonDioxCalc = energyCalc.sheet_by_index(0)
    global COTotal

    for row_num2 in range(carbonDioxCalc.nrows):
        row_value2=carbonDioxCalc.row_values(row_num2)
        if row_value2[1]==eGRID:
            global COtotal
            COtotal=row_value2[6]
            global kwhtotal
            kwhtotal=row_value2[5]

    timeSheet = xlrd.open_workbook('/projects/b1045/EVTool/'+eGRID+'.xlsx')
    sheetTwo = timeSheet.sheet_by_index(0)
    #theirStart = int(startEnter.get())
    #theirEnd = int(endEnter.get())
    theirStart = 2018
    theirEnd = 2051

    while True:
        for startRow in range(sheetTwo.nrows):
            startDateRow=sheetTwo.row_values(startRow)
            if startDateRow[1]==theirStart:
                yearFuture = sheetTwo.row_values(rowx=-theirStart + 2052)
                try:
                    COpercentage =((startDateRow[2] * 95.6191)/(yearFuture[2] * 95.6191))
                except ZeroDivisionError:
                    COpercentage=1

                kwhpercentage = startDateRow[6]/yearFuture[6]
                COtotal = COtotal * COpercentage
                kwhtotal = kwhtotal * kwhpercentage
                factor = COtotal/kwhtotal
                if v==1:
                    coalList.append(factor*18724.27*0.1988*2.205)
                else:
                    coalList.append(factor*int(theirnumm)*.1988*2.205*1.60934)



        theirStart+=1
        if theirStart>=theirEnd:
            break

    return coalList

def getNG():
    NGList=[]
    zipGrids = xlrd.open_workbook('/projects/b1045/EVTool/zipGrid.xlsx')
    sheet = zipGrids.sheet_by_index(4)

    for row_num in range(sheet.nrows):#sort thru all rows in excel
        row_value = sheet.row_values(row_num)#sorts thru rows
        if row_value[1] == int(enter):#cycles until input equals a row
            global eGRID
            eGRID = row_value[3]

    energyCalc = xlrd.open_workbook('/projects/b1045/EVTool/energyCalculations.xlsx')
    carbonDioxCalc = energyCalc.sheet_by_index(0)
    global COTotal

    for row_num2 in range(carbonDioxCalc.nrows):
        row_value2=carbonDioxCalc.row_values(row_num2)
        if row_value2[1]==eGRID:
            global COtotal
            COtotal=row_value2[7]
            global kwhtotal
            kwhtotal=row_value2[5]

    timeSheet = xlrd.open_workbook('/projects/b1045/EVTool/'+eGRID+'.xlsx')
    sheetTwo = timeSheet.sheet_by_index(0)
    #theirStart = int(startEnter.get())
    #theirEnd = int(endEnter.get())
    theirStart = 2018
    theirEnd = 2051

    while True:
        for startRow in range(sheetTwo.nrows):
            startDateRow=sheetTwo.row_values(startRow)
            if startDateRow[1]==theirStart:
                yearFuture = sheetTwo.row_values(rowx=-theirStart + 2052)
                try:
                    COpercentage =((startDateRow[4] * 53.06)/(yearFuture[4] * 53.06))
                except ZeroDivisionError:
                    COpercentage=1
                kwhpercentage = startDateRow[6]/yearFuture[6]
                COtotal = COtotal * COpercentage
                kwhtotal = kwhtotal * kwhpercentage
                factor = COtotal/kwhtotal
                if v==1:
                    NGList.append(factor*18724.27*0.1988*2.205)
                else:
                    NGList.append(factor*int(theirnumm)*.1988*2.205*1.60934)


        theirStart+=1
        if theirStart>=theirEnd:
            break

    return NGList

def getOil():
    OilList=[]
    zipGrids = xlrd.open_workbook('/projects/b1045/EVTool/zipGrid.xlsx')
    sheet = zipGrids.sheet_by_index(4)

    for row_num in range(sheet.nrows):#sort thru all rows in excel
        row_value = sheet.row_values(row_num)#sorts thru rows
        if row_value[1] == int(enter):#cycles until input equals a row
            global eGRID
            eGRID = row_value[3]

    energyCalc = xlrd.open_workbook('/projects/b1045/EVTool/energyCalculations.xlsx')
    carbonDioxCalc = energyCalc.sheet_by_index(0)
    global COTotal

    for row_num2 in range(carbonDioxCalc.nrows):
        row_value2=carbonDioxCalc.row_values(row_num2)
        if row_value2[1]==eGRID:
            global COtotal
            COtotal=row_value2[8]
            global kwhtotal
            kwhtotal=row_value2[5]

    timeSheet = xlrd.open_workbook('/projects/b1045/EVTool/'+eGRID+'.xlsx')
    sheetTwo = timeSheet.sheet_by_index(0)
    #theirStart = int(startEnter.get())
    #theirEnd = int(endEnter.get())
    theirStart = 2018
    theirEnd = 2051

    while True:
        for startRow in range(sheetTwo.nrows):
            startDateRow=sheetTwo.row_values(startRow)
            if startDateRow[1]==theirStart:
                yearFuture = sheetTwo.row_values(rowx=-theirStart + 2052)
                try:
                    COpercentage =(startDateRow[4] * 53.06)/(yearFuture[4] * 53.06)
                except ZeroDivisionError:
                    COpercentage=1

                kwhpercentage = startDateRow[6]/yearFuture[6]
                COtotal = COtotal * COpercentage
                kwhtotal = kwhtotal * kwhpercentage
                factor = COtotal/kwhtotal
                if v==1:
                    OilList.append(factor*18724.27*0.1988*2.205)
                else:
                    OilList.append(factor*int(theirnumm)*.1988*2.205*1.60934)

        theirStart+=1
        if theirStart>=theirEnd:
            break

    return OilList

def getBio():
    BioList=[]
    zipGrids = xlrd.open_workbook('/projects/b1045/EVTool/zipGrid.xlsx')
    sheet = zipGrids.sheet_by_index(4)

    for row_num in range(sheet.nrows):#sort thru all rows in excel
        row_value = sheet.row_values(row_num)#sorts thru rows
        if row_value[1] == int(enter):#cycles until input equals a row
            global eGRID
            eGRID = row_value[3]

    energyCalc = xlrd.open_workbook('/projects/b1045/EVTool/energyCalculations.xlsx')
    carbonDioxCalc = energyCalc.sheet_by_index(0)
    global COTotal

    for row_num2 in range(carbonDioxCalc.nrows):
        row_value2=carbonDioxCalc.row_values(row_num2)
        if row_value2[1]==eGRID:
            global COtotal
            COtotal=row_value2[9]
            global kwhtotal
            kwhtotal=row_value2[5]

    timeSheet = xlrd.open_workbook('/projects/b1045/EVTool/'+eGRID+'.xlsx')
    sheetTwo = timeSheet.sheet_by_index(0)
    #theirStart = int(startEnter.get())
    #theirEnd = int(endEnter.get())
    theirStart = 2018
    theirEnd = 2051

    while True:
        for startRow in range(sheetTwo.nrows):
            startDateRow=sheetTwo.row_values(startRow)
            if startDateRow[1]==theirStart:
                yearFuture = sheetTwo.row_values(rowx=-theirStart + 2052)
                try:
                    COpercentage =((.1126 * startDateRow[5] * 52.04))/((.1126 * yearFuture[5] * 52.04))
                except ZeroDivisionError:
                    COpercentage=1
                kwhpercentage = startDateRow[6]/yearFuture[6]
                COtotal = COtotal * COpercentage
                kwhtotal = kwhtotal * kwhpercentage
                factor = COtotal/kwhtotal
                if v==1:
                    BioList.append(factor*18724.27*0.1988*2.205)
                else:
                    BioList.append(factor*int(theirnumm)*.1988*2.205*1.60934)


        theirStart+=1
        if theirStart>=theirEnd:
            break

    return BioList



valueList = getgrid()
coalVal = getCoal()
NGVal = getNG()
OilVal=getOil()
BioVal = getBio()

totalList = []
counter = 0
while counter<=32:
    totalList.append(coalVal[counter]+NGVal[counter]+OilVal[counter]+BioVal[counter])
    counter+=1


coalCounter = 0
coalPercent = []
while coalCounter<=32:
    try:
        coalPercent.append(coalVal[coalCounter]/totalList[coalCounter])
    except ZeroDivisionError:
        coalPercent.append(0)
    coalCounter+=1


totalCoalCounter=0
totalCoal = []
while totalCoalCounter<=32:
    try:
        totalCoal.append(coalPercent[totalCoalCounter]*valueList[totalCoalCounter])
    except ZeroDivisionError:
        totalCoal.append(0)
    totalCoalCounter+=1


NGCounter = 0
NGPercent = []
while NGCounter <= 32:
    try:
        NGPercent.append(NGVal[NGCounter] / totalList[NGCounter])
    except ZeroDivisionError:
        NGPercent.append(0)
    NGCounter += 1


totalNGCounter = 0
totalNG = []
while totalNGCounter <= 32:
    try:
        totalNG.append(NGPercent[totalNGCounter] * valueList[totalNGCounter])
    except ZeroDivisionError:
        totalNG.append(0)
    totalNGCounter += 1


oilCounter = 0
oilPercent = []
while oilCounter <= 32:
    try:
        oilPercent.append(OilVal[oilCounter] / totalList[oilCounter])
    except ZeroDivisionError:
        oilPercent.append(0)
    oilCounter += 1


totalOilCounter = 0
totalOil = []
while totalOilCounter <= 32:
    totalOil.append(oilPercent[totalOilCounter] * valueList[totalOilCounter])
    totalOilCounter += 1


bioCounter = 0
bioPercent = []
while bioCounter <= 32:
    try:
        bioPercent.append(BioVal[bioCounter] / totalList[bioCounter])
    except ZeroDivisionError:
        bioPercent.append(0)
    bioCounter += 1


totalBioCounter = 0
totalBio = []
while totalBioCounter <= 32:
    totalBio.append(bioPercent[totalBioCounter] * valueList[totalBioCounter])
    totalBioCounter += 1




startcounter = int(startyear)
yearlist = []
while startcounter < int(endyear):
    yearlist.append(startcounter)
    startcounter += 1


final = np.sum(valueList[int(startyear)-2017:int(endyear)-2017])

print('the total amount of CO2 emissions is '+str(final)+' pounds')
average = final/(int(endyear)-int(startyear))
print('The average per year is: '+str(average)+' pounds')

plt.plot(yearlist, valueList[int(startyear)-2017:int(endyear)-2017])
plt.plot(yearlist, totalCoal[int(startyear)-2017:int(endyear)-2017])
plt.plot(yearlist, totalNG[int(startyear)-2017:int(endyear)-2017])
plt.plot(yearlist, totalOil[int(startyear)-2017:int(endyear)-2017])
plt.plot(yearlist, totalBio[int(startyear)-2017:int(endyear)-2017])
plt.legend(["Total","Coal","Natural Gas", 'Oil', 'Renewables'])
plt.xlabel('Year')
plt.ylabel('Emissions (eqiv. pounds of CO2)')
plt.title('CO2 Emissions Over Time in '+eGRID)
plt.show()

#def buttonPrint():
    #print(valueList)
#b = tk.Button(root, text="Run", command=buttonPrint)
#b.grid(row=3, column=0)



#root.mainloop()
