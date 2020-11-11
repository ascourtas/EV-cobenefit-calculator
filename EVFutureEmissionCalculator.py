import xlrd  #allows for reading from excel
import matplotlib.pyplot as plt  #allows for creating figures within python
import numpy as np  #allows for higher level math functions
import seaborn as sns  #makes figures more organized
sns.set()  #apply seaborn basics

AZNM = xlrd.open_workbook('/projects/b1045/EVTool/AZNM.xlsx')  #files with future data for each type of power generation by egrid
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

enter = input('what is your zip code? ')  #prompts with these questions for inputs to be used in calculations
startyear = input('what year do you plan on starting to own the vehicle? ')
endyear = input('when do you plan to stop owning the vehicle? ')

global startEnter
global endEnter


def getgrid():  #defines function

    vList = []  #opens list to be appended to later
    zipGrids = xlrd.open_workbook('/projects/b1045/EVTool/zipGrid.xlsx')  #defines variable for a file with all zip codes and grids
    sheet = zipGrids.sheet_by_index(4)  #chooses the propper sheet within the excel file with the data we want
    for row_num in range(sheet.nrows):  #for a particular row in the set of all zip codes in the continental us
       row_value = sheet.row_values(row_num)  #narrows search down to row by row
       if row_value[1] == int(enter):  #if the value of the first column in a given row is equal to zip code input
        global eGRID
        eGRID = row_value[3]  #variable eGRID is set for the third column value in that row
    energyCalc = xlrd.open_workbook('/projects/b1045/EVTool/energyCalculations.xlsx')  #defines variable for file with total co2
                                                                            #and kWh in every grid
    carbonDioxCalc = energyCalc.sheet_by_index(0)  #chooses correct sheet within excel file
    global COTotal
    for row_num2 in range(carbonDioxCalc.nrows):  #for a particular row in the energy calc excel sheet
        row_value2 = carbonDioxCalc.row_values(row_num2)  #narrows down to particular rows again
        if row_value2[1] == eGRID:  #if the value of the first column equals the grid from the zipGrids file
            global COtotal
            COtotal = row_value2[4]  #total CO2 of this grid is the 5th column
            global kwhtotal
            kwhtotal=row_value2[5]  #total kWh of this grid is the 6th column
    timeSheet = xlrd.open_workbook('/projects/b1045/EVTool/'+eGRID+'.xlsx')  #open the excel file that corresponds to the correct grid
    sheetTwo = timeSheet.sheet_by_index(0)  #chooses correct sheet in the eGRID excel file
    theirStart = 2018  #these variables set the extremes of the years possible for our function to consider
    theirEnd = 2051

    while True:  #sets the boundaries of the calculation
        for startRow in range(sheetTwo.nrows):
            startDateRow = sheetTwo.row_values(startRow)  #
            if startDateRow[1] == theirStart:
                yearFuture = sheetTwo.row_values(rowx=-theirStart + 2052)
                global COpercentage
                COpercentage = ((startDateRow[2] * 95.6191) + (startDateRow[4] * 53.06) + (startDateRow[3] * 74.57193) + (.1126 * startDateRow[5] * 52.04))/((yearFuture[2] * 95.6191) + (yearFuture[4] * 53.06) + (yearFuture[3] * 74.57193) + (.1126 * yearFuture[5] * 52.04))
                kwhpercentage = startDateRow[6]/yearFuture[6]
                COtotal = COtotal * COpercentage
                kwhtotal = kwhtotal * kwhpercentage
                factor = COtotal/kwhtotal
                kmyear = (18724.27 * 0.995 ** (theirStart-2019))
                vList.append(factor*kmyear*0.1988*2.205)

        theirStart += 1
        if theirStart >= theirEnd:
            break

    return vList


valueList = getgrid()


finalelectric = np.sum(valueList[int(startyear)-2017:int(endyear)-2017])
finalhybrid = 6258 * (int(endyear)-int(startyear))
finalicev = 11435 * (int(endyear)-int(startyear))
finalplugin = ((finalelectric*.55*1.146875)+(.45*.6412*finalicev))

x_pos = ['Electric Vehicle', 'Hybrid Plug-in', 'Hybrid Gasoline', 'Gasoline Vehicle']
y_value = [finalelectric, finalplugin, finalhybrid, finalicev]

plt.bar(x_pos, y_value, color=(0, 0.38, 0.11, 1))
plt.xlabel("Car Type")
plt.ylabel("lbs of CO2 Equivalent")
plt.title("Total CO2 Emissions by Car Type in eGrid: " + eGRID)
plt.xticks(x_pos, x_pos)


plt.show()


print('the total amount of CO2 emissions from an electric car is '+str(finalelectric)+' pounds')
average = finalelectric/(int(endyear)-int(startyear))
print('The average per year is: '+str(average))
