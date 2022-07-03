from computeGradient import computeGradient
import math

def heightAtGradMPt(h1,h2,spatialRes):
    """
    Computes height at the point from which the gradient measurement was taken

    Args:

    h1 : elevation 1

    h2 : elevation 2

    spatialRes : spatial resolution (metres squared)


    Returns:
    """
    grad = computeGradient(h1,h2,spatialRes)
    height = math.tan(math.radians(grad)*math.sqrt(((spatialRes/2)**2)*2))+min(h1,h2)
    return {
        'grad': grad,
        'height': height
    }
    