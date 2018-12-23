from csvReader import *
from tableCreator import *
from math import *
import string

class PlaneManager():

    def __init__(self, csvName, xlsName):
        self.__csvName=csvName
        self.__xlsName=xlsName
        self.__discDict={'Disc':[],'Type':[],'Plan':[]}
        self.__distTypes=('Лекция','Лаб. работа', 'Практич. Занятие', 'Зачет', 'Экзамен', 'Диф. Зачет')
        self.__reader=csvReader(csvName)
        self.__table=XLSTable(xlsName)
        self.__countFieldinPlan=0
        self.__groupCount=0
        self.__groups_col={}
        self.__groupRowConst=3
        self.__columnSymbolEqualsNumber=list(string.ascii_uppercase)

    def createTemplate(self):
        self.__table.initFillTables()

    def addGroups(self, groupNameList):
        self.__table.writeRectDataInTable([groupNameList],3,self.__groupCount+4,0)
        for i in range(0,len(groupNameList)):
            self.__groups_col[groupNameList[i]]=i+self.__groupCount+4
        self.__groupCount+=len(groupNameList)


    def __addPlanPoint(self,discName, type, plan):
        self.__discDict['Disc'].append(discName)
        self.__discDict['Type'].append(type)
        self.__discDict['Plan'].append(plan)

    def getPlanPointOfSCV(self):
        data=self.__reader.csvReadInList()
        print("Add lessons in plan:")
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
        self.__table.autoFitDimmensionsCollumn(0,list(map(lambda x: self.__columnSymbolEqualsNumber[x-1],list(self.__groups_col.values()))))
        self.__table.autoFitDimmensionsCollumn(1)
        self.__table.autoFitDimmensionsCollumn(2)

    def addFormulsforLessonsinPlanOfList(self):
        print("Добавленно предметов в план: " + str(self.__countFieldinPlan))
        grouplist=list(self.__groups_col.keys())
        for g in grouplist:
            ncol=self.__groups_col[g]
            scol=self.__columnSymbolEqualsNumber[ncol-1]
            cellIndexForGroupName=str(scol)+"$"+str(self.__groupRowConst)
            for j in range(self.__groupRowConst+1,self.__groupRowConst+self.__countFieldinPlan+1):
                formuls="=2*COUNTIFS(Расписание!$E$2:$E$10000,\'Сводная страница\'!$A" +str(j)+",Расписание!$I$2:$I$10000,\'Сводная страница\'!"+cellIndexForGroupName+",Расписание!$F$2:$F$10000,\'Сводная страница\'!$B"+str(j)
                dataCell=[[formuls]]
                self.__table.writeRectDataInTable(dataCell,j,ncol,0)
    def getGroupColumn(self, groupName):
        try:
            col=self.__groups_col[groupName]
        except KeyError:
            print("Not found in table group", groupName)
            exit()
        return col
