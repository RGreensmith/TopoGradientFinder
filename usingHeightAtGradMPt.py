from heightAtGradMeasurePt import heightAtGradMPt
from computeGradient import computeGradient

spatialRes = 50
a = heightAtGradMPt(20,35,spatialRes)
b = heightAtGradMPt(75,100,spatialRes)

print("max a,b = ", max(a['height'],b['height'])," a = ",a," b = ",b)

maxHeight = max(a['height'],b['height'])
if (a['height'] == maxHeight):
    print("aGrad = ",a['grad'])

if (b['height'] == maxHeight):
    print("bGrad = ",b['grad'])