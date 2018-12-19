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
groups=argv[3:]

print("Add groups: "+ str(groups))

filenamecsv=getAbsolutePathforFile(checkFileExtension(argv[1],".csv"))
filenamexlsx=getAbsolutePathforFile(checkFileExtension(argv[2],".xlsx"))

planeM=PlaneManager(filenamecsv,filenamexlsx)
planeM.createTemplate()
planeM.addGroups(groups)
planeM.getPlanPointOfSCV()
planeM.writePlanPointsinTable()
planeM.makeHumanReadable()
planeM.savePlane()
