import xlrd
import xlwt
import json

mainObj = {}
headLines = []
outputRows = []
loc = "/Users/joel/Documents/DataFiles/HistoryOMX.xls"
outLoc = "/Users/joel/Documents/DataFiles/HistoryOMXformatted.xlsx"

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
print(sheet)
numCols = sheet.ncols
nRows = sheet.nrows


def looper():
    rowList = []
    for i in range(nRows):
        if i != 0:
            for j in range(len(sheet.row(i))-4):
                rowList.append(sheet.cell_value(i, j))
            recAdd(rowList, mainObj)
            rowList = []
        else:
            for j in range(len(sheet.row(i))-1):
                headLines.append(sheet.cell_value(i, j))
    mainObj['HeadLines'] = headLines


def recAdd(data, obj):
    if len(data) > 0:
        if data[0] in obj:
            sendObj = obj[data[0]]
            data.pop(0)
            recAdd(data, sendObj)
        else:
            obj[data[0]] = {}
            sendObj = obj[data[0]]
            data.pop(0)
            recAdd(data, sendObj)
    
    return

def prepForSum(myDict):
    del myDict["HeadLines"]
    myDict = sumLastRow(myDict, [])
    myDict["HeadLines"] = headLines
    return myDict


def sumLastRow(myDict, keys):
    if(len(myDict) != 0):
        for key in myDict.keys():
            keys.append(key)
            
            if(len(myDict[key]) == 0):
                myDict = sum(myDict.keys())/len(myDict.keys())
                return myDict
            else:
                myDict[key] = sumLastRow(myDict[key], keys)
        return myDict

def concatRowVector(obj, values):
    for key in obj.keys():
        myValues = [] + values
        if(key != "HeadLines"):
            if type(obj[key]) is dict:
                myValues.append(key)
                concatRowVector(obj[key], myValues)
            else:
                valVec = myValues + [key, obj[key]]
                outputRows.append(valVec)

def writeToExcel():
    outWb = xlwt.Workbook(encoding = 'ascii')
    worksheet = outWb.add_sheet('Sheet')
    for i in range(len(outputRows)):
        for j in range(len(outputRows[i])):
            worksheet.write(i, j, label = outputRows[i][j])
    outWb.save(outLoc)



#Calc length of the lowest tier. May not be used
def calcLength(obj):
    for key in obj.keys():
        if(key != "HeadLines"):
            if type(obj[key]) is dict:
                length = calcLength(obj[key])
            else:
                length = len(obj)
    return length


looper()
maiObj = prepForSum(mainObj)

concatRowVector(mainObj, [])

if True:
    headLines = ['Year', 'Month', 'Cost']
    mainObj["HeadLines"] = headLines
    outputRows.insert(0, headLines)

writeToExcel()
print("Finish")