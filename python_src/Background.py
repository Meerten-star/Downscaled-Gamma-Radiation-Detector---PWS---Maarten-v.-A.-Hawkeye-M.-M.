import matplotlib.pyplot as plt


def getPoints():
    file = open("Spectra/BackgroundEstimatedPoints.txt", "r")
    dataPoints = []

    for line in file:
        x, y = line.strip().split()
        x = float(x[:-1])
        y = round(float(y))
        dataPoints.append((x, y))

    return dataPoints

calibration = (-7.703226, 1.025806) # start, step
dataPoints = getPoints()
spectrumData = []
energyList = []

channel = 0
energy = calibration[0]
pointIndex = 0
while energy < 734:
    r = input(f"\nenergy: {energy}. (energy+1: {energy + calibration[1]}). ch{channel}"
              f"\npoint {dataPoints[pointIndex]}. next point {dataPoints[pointIndex + 1]}\n")
    if r == "1":
        pointIndex += 0
        spectrumData.append(dataPoints[pointIndex][1])
    elif r == "2":
        pointIndex += 1
        spectrumData.append(dataPoints[pointIndex][1])
    else: # accident / other self input
        r = input("new count: ")
        spectrumData.append(int(r))

    print(spectrumData[-3:])

    energyList.append(energy)

    channel += 1
    energy = calibration[0] + calibration[1] * channel

print(energyList)
print(spectrumData)

plt.plot(energyList, spectrumData)
plt.show()


