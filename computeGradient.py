import math

def computeGradient(h1,h2,spatialRes):
    """
    Computes topographic gradient

    Args:

    h1 = elevation 1

    h2 = elevation 2

    spatialRes = spatial resolution (metres squared)


    Returns: inverse tan of theta (unit = degrees)
    """
    lat1 = 5
    lat2 = lat1+spatialRes
    lon1 = 35
    lon2 = lon1+spatialRes

    opposite = abs(h1-h2)
    adjacent = abs(lat1-lat2)**2+abs(lon1-lon2)**2
    theta = math.degrees(math.atan(opposite/adjacent))

    print ("Theta = ",theta)
    return theta