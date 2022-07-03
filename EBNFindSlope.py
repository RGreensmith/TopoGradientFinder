import pandas as pd
from EBNProjSlopes import EBNProjSlopes
# from loader import loader
from qgis.PyQt.QtCore import *
from qgis.core import *
from qgis import processing
from pyQGIS import slopeInsidePoly
from OSNationalGridFinder import OSNationalGridFinder

# Supply path to qgis install location
QgsApplication.setPrefixPath("/path/to/qgis/installation", True)

# Create a reference to the QgsApplication. Setting the second arg to False disables the GUI
qgs = QgsApplication([], False)

# Load providers
qgs.initQgis()

RLBFileNames = pd.read_csv('to_ignore/RLBFileNames.csv',header=0)
pathToOSTerrAsciis = ''
OSTerrAscii = OSNationalGridFinder(RLBFileNames,pathToOSTerrAsciis)
EBNProjSlopes(RLBFileNames,OSTerrAscii)

# exitQgis removes the provider and layer registries from memory
qgs.exitQgis()