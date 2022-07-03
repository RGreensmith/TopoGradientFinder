from pandas import pd
from eastingNorthingToOSNG import eastingNorthingToOSNG
from getShpEastingNorthing import getShpEastingNorthing
def OSNationalGridFinder(RLBFileNames,pathToRLBs,pathToOSTerrAsciis):
    
    refDF= {}
    refDF = ["RLBFileNames"]

    for file in RLBFileNames:
        #get eastings and northings from RLB shapefile
        RLBEastNor = getShpEastingNorthing(pathToRLBs,file["RLBFileNm"])
        # EastNorToOSNG
        eastingNorthingToOSNG(RLBEastNor["easting"],RLBEastNor["northing"])

    return refDF