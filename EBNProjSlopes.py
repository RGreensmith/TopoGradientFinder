import pandas as pd
def EBNProjSlopes(RLBFileNames,OSTerrAscii):

    projSlopesDict = {}
    projSlopesDict = ["maxSlope"]
    projSlopesDict = ["maxSlope"]
    projSlopesDict = ["maxSlope"]

    for a in EBNProjs:
        # load raster
        # rlayer = QgsProject.instance("NVC.qgz").mapLayersByName('su02_OST50GRID_20220506 — SU02.asc')[0]
        pathway = 'c:/path/to/raster/file.tif'
        rLyr = iface.addRasterLayer(pathway, 'layer name')

        # load shapefile
        ldr = loader(qgis.utils.iface)
        RLB = ldr.load_shapefiles('/my/path/to/shapefile/directory')

        maxSlope = slopeInsideRLB(rLyr,RLB,outPath)

        projSlopesDict[a] = RLBFileNames[1,a]
        projSlopesDict[a] = RLBFileNames[2,a]
        projSlopesDict[maxSlope][a] = maxSlope
    
    return projSlopesDict