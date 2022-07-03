from qgis.PyQt.QtCore import *
from qgis.core import *
from qgis import processing

from slopeGdalWrapper import slopeGdalWrapper
from slopeQGISWrapper import slopeQGISWrapper

def slopeInsidePoly(rLyr,RLB,outPath):

    # clip raster by mask layer
    rLyrOSTerr50 = processing.runalg('gdalogr:cliprasterbymasklayer', rLyr, rLB, no_data, alpha_band, keep_resolution, extra, output)

    # create slope raster
    outGIS = ''
    outGdal = ''
    rLyrOSTerr50SlopeQGIS = slopeQGISWrapper(rLyrOSTerr50,outGIS)
    rLyrOSTerr50SlopeGDAL = slopeGdalWrapper(rLyrOSTerr50,outGdal)

    # find max slope
    rLyrOSTerr50Stats = rlayer.dataProvider().bandStatistics(1, QgsRasterBandStats.All)
    maxSlope = rLyrOSTerr50Stats.maximumValue
    print("maxSlope = ",maxSlope)
    return maxSlope