from PlaneManager import *
import sys
import re
import os

def getAbsolutePathforFile(name):
    if name[0]=="/":
        return name
    else:
        return os.path.dirname(os.path.abspath(__file__))+"/../../"+name

def checkFileExtension(name, extension):
    if re.search(extension,name):
        filename=name
    else:
        filename=name+extension
    return filename

argv=sys.argv

flag=False
curflag=""
groups=[]
for a in argv:
    if a[0]=="-" or a[0]=="--":
        flag=True
        curflag=a
        continue
    if curflag=="-g" or curflag=="--groups":
        groups.append(a)



print("Add groups: "+ str(groups))

filenamecsv=getAbsolutePathforFile(checkFileExtension(argv[1],".csv"))
filenamexlsx=getAbsolutePathforFile(checkFileExtension(argv[2],".xlsx"))
planeM=PlaneManager(filenamecsv,filenamexlsx)
planeM.createTemplate()
planeM.addGroups(groups)
planeM.getPlanPointOfSCV()
planeM.writePlanPointsinTable()
planeM.addFormulsforLessonsinPlanOfList()
planeM.makeHumanReadable()
planeM.savePlane()
