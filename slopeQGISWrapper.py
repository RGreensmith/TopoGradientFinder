from sqlalchemy import true
from qgis.PyQt.QtCore import *
from qgis.core import *
from qgis import processing

def slopeQGISWrapper(inPath,outPath = "",zfact = 1.0,add_iface = true):
    processing.run(
        "qgis:slope", \
        {
            'INPUT':inPath, \
            'Z_FACTOR':zfact, \
            'OUTPUT':outPath
        }
    )
    if add_iface:
            iface.addRasterLayer(outPath)