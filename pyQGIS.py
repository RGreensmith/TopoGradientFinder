from qgis.PyQt.QtCore import *
from qgis.core import *
from qgis import processing
from loader import Loader
from slopeGdalWrapper import slopeGdalWrapper
from slopeQGISWrapper import slopeQGISWrapper

# Supply path to qgis install location
QgsApplication.setPrefixPath("/path/to/qgis/installation", True)

# Create a reference to the QgsApplication.  Setting the
# second argument to False disables the GUI.
qgs = QgsApplication([], False)

# Load providers
qgs.initQgis()

# Write your code here to load some layers, use processing
# algorithms, etc.

# Finally, exitQgis() is called to remove the
# provider and layer registries from memory
qgs.exitQgis()

# load raster
#rlayer = QgsProject.instance("NVC.qgz").mapLayersByName('su02_OST50GRID_20220506 — SU02.asc')[0]
pathway = 'c:/path/to/raster/file.tif'
rLyr = iface.addRasterLayer(pathway, 'layer name')

# load shapefile
ldr = Loader(qgis.utils.iface)
rLB = ldr.load_shapefiles('/my/path/to/shapefile/directory')

# clip raster by mask layer
rLyrOSTerr50 = processing.runalg('gdalogr:cliprasterbymasklayer', rLyr, rLB, no_data, alpha_band, keep_resolution, extra, output)

# create slope raster
outGIS = ''
outGdal = ''
slopeQGISWrapper(rLyrOSTerr50,outGIS)
slopeGdalWrapper(rLyrOSTerr50,outGdal)

# find max slope
rLyrOSTerr50Stats = rlayer.dataProvider().bandStatistics(1, QgsRasterBandStats.All)
maxSlope = rLyrOSTerr50Stats.maximumValue
print("maxSlope = ",maxSlope)