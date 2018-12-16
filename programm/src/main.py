from PlaneManager import *
import sys


argv=sys.argv
filenamecsv=argv[1]
filenamexls=argv[2]
planeM=PlaneManager(filenamecsv,filenamexls)
planeM.createTemplate()
planeM.getPlanPointOfSCV()
planeM.writePlanPointsinTable()
planeM.makeHumanReadable()
planeM.savePlane()
