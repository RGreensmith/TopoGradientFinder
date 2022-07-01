import math
h1=1
h2=2
opposite = abs(h1-h2)
lat1 = 5
lat2 = 55
lon1 = 35
lon2 = 85
adjacent = abs(lat1-lat2)**2+abs(lon1-lon2)**2
theta = math.degrees(math.atan(opposite/adjacent))