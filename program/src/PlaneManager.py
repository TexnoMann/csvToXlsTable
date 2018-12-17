from csvReader import *
from tableCreator import *
from math import *

class PlaneManager():

    def __init__(self, csvName, xlsName):
        self.__csvName=csvName
        self.__xlsName=xlsName
        self.__discDict={'Disc':[],'Type':[],'Plan':[]}
        self.__distTypes=('Лекция','Лаб. работа', 'Практич. Занятие', 'Зачет', 'Экзамен', 'Диф. Зачет')
        self.__reader=csvReader(csvName)
        self.__table=XLSTable(xlsName)
        self.__countFieldinPlan=0

    def createTemplate(self):
        self.__table.initFillTables()

    def __addNewGroup(self, groupName):
        if not self__discDict[groupName]:
            self.__discDict[groupName]=[]

    def __addPlanPoint(self,discName, type, plan):
        self.__discDict['Disc'].append(discName)
        self.__discDict['Type'].append(type)
        self.__discDict['Plan'].append(plan)

    def getPlanPointOfSCV(self):
        print("Add lessons in plan:")
        data=self.__reader.csvReadInList()
        for i in range(0,len(data)):
            if not(data[i][2] in self.__discDict['Disc']):
                for j in range(3,6):
                    if  not isnan(data[i][j]):
                        self.__addPlanPoint(data[i][1],self.__distTypes[j-3],data[i][j])
                        print([data[i][1],self.__distTypes[j-3],data[i][j]])
                        self.__countFieldinPlan+=1
                hours=0;
                if(data[i][7]=='Экзамен'): hours=4
                else: hours=2
                if hours !=0:
                    self.__addPlanPoint(data[i][1],data[i][7],hours)
                    print([data[i][1],data[i][7],hours])
                    self.__countFieldinPlan+=1

    def writePlanPointsinTable(self):
        for i in range(0,self.__countFieldinPlan):
            field=[[self.__discDict['Disc'][i],self.__discDict['Type'][i],self.__discDict['Plan'][i]]]
            place =i+4
            self.__table.writeRectDataInTable(field,place,1,0)

    def savePlane(self):
        self.__table.saveInFile()

    def makeHumanReadable(self):
        self.__table.autoFitDimmensionsCollumn(0)
        self.__table.autoFitDimmensionsCollumn(1)
        self.__table.autoFitDimmensionsCollumn(2)
