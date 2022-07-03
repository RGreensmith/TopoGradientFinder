from sqlalchemy import false, true
from qgis.PyQt.QtCore import *
from qgis.core import *
from qgis import processing

def slopeGdalWrapper(inPath,outPath = "",percent = true):
    processing.run(
        "gdal:slope", \
        {
            'Input':inPath, \
            'BAND':1, \
            'SCALE':1, \
            'As_Percent':percent, \
            'COMPUTE_EDGES':false, \
            'ZEVENBERGEN':false, \
            'OPTIONS':'', \
            'OUTPUT': outPath
        }
    )
    iface.addRasterLayer(outPath)