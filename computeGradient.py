import math

def computeGradient(h1,h2,spatialRes):
    """
    Computes topographic gradient

    Args:

    h1 : elevation 1

    h2 : elevation 2

    spatialRes : spatial resolution (metres squared)


    Returns: inverse tan of theta (unit = degrees)
    """
    opposite = abs(h1-h2)
    adjacent = math.sqrt((spatialRes**2)*2)
    theta = math.degrees(math.atan(opposite/adjacent))

    print ("Theta = ",theta," Adjacent = ",adjacent," Opposite = ",opposite)
    return theta