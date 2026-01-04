import matplotlib.pyplot as plt
from matplotlib.widgets import SpanSelector
import numpy as np


def getSpectrum(file):
    info = [] # live time, real time, start value, step value
    data = []
    file = open(file, "r")
    line = file.readline().strip() # .strip() removes \n
    state = "find info"
    while line:
        line = line.strip()
        if state == "find info":
            if "$MEAS_TIM:" in line or "$SPEC_CAL:" in line:
                infoPoints = file.readline().strip().split()[:2]
                for s in infoPoints:
                    info.append(float(s))
            elif "$DATA:" in line:
                file.readline()
                state = "get data"
        elif state == "get data":
            dataPoints = line.split()
            dataPoints = [int(s) for s in dataPoints]
            data += dataPoints
        line = file.readline().strip()
    file.close()
    return info, data

def subtractBackground(sourceInfo, sourceData):
    backgroundInfo, backgroundData = getSpectrum("Spectra/EstimatedBackgroundSpectrum.spe")
    timeFactor = sourceInfo[0] / backgroundInfo[0] # liveTime
    data = []
    for i in range(len(sourceData)):
        dataPoint = sourceData[i] - backgroundData[i] * timeFactor
        if dataPoint < 0:
            dataPoint = 0
        data.append(dataPoint)
    return data


def displayAndSelect(info, data, data_notSub):
    class Selection:
        def __init__(self):
            self.selectionDomain = ()

        def onSelect(self, xMin, xMax):
            self.selectionDomain = (xMin, xMax)
            print(f"Selection: {xMin} - {xMax}")

    fig, ax = plt.subplots()

    # Spectrum data
    startValue, stepValue = info[2:]
    x = [i * stepValue + startValue for i in range(len(data))]
    y = data_notSub

    # Background data
    bground_info, bground_data = getSpectrum("Spectra/EstimatedBackgroundSpectrum.spe")
    bground_startValue, bground_stepValue = bground_info[2:]
    x_bground = [i * bground_stepValue + bground_startValue for i in range(len(bground_data))]
    y_bground = bground_data

    # Background subtracted spectrum data
    y_sub = data

    # make plots
    ax.plot(x, y)
    ax.plot(x_bground, y_bground)
    ax.plot(x, y_sub)

    # define selection
    selector = Selection()
    span = SpanSelector(ax, onselect=selector.onSelect, direction="horizontal", useblit=True, props=dict(alpha=0.3, facecolor="tab:blue"), interactive=True, drag_from_anywhere=True)

    # shows the plot, user can now make selection, continues after closing screen.
    plt.show()

    # returns the selection domain
    return selector.selectionDomain

def energyToChannel(info, energy):
    startValue, stepValue = info[2:]
    channel = (energy - startValue) / stepValue
    return int(channel)

def getSelectionCPS(info, data, selectionDomain): # in channel space
    # get selected part of the data
    selection = data[energyToChannel(info, selectionDomain[0]) : energyToChannel(info, selectionDomain[0]) + 1] # both inclusive

    # calculate the sum of the selection
    totalCounts = sum(selection)

    # correct for live time (CPS is counts/s)
    totalCPS = totalCounts / info[0]
    return totalCPS


def recalculateFileSelectionsCPS():
    selectionDomains = []
    file = open("selections.txt", "r")
    for line in file:
        if "Selection from" in line:
            splitLine = line.strip().split()
            xMin = float(splitLine[4][1:-1])
            xMax = float(splitLine[9][1:])
            selectionDomains.append((xMin, xMax))
    return selectionDomains


def getActivityPerIsotope(request, infoList, dataList, dataList_notSubtracted, emissionPeaks, detectorEfficiency):
    activityPerIsotope = []

    n = len(emissionPeaks)
    if request == "recalculate":
        selectionDomains = recalculateFileSelectionsCPS()
        if n != len(selectionDomains):
            n = len(selectionDomains)
            print("warning, no consistent peak attribution")

    for i in range(n):
        if request == "new selection":
            print(f"Select peak {i + 1}/{n} at: {emissionPeaks[i][1]} keV of isotope {emissionPeaks[i][0]}")

            # Let the user make a selection (shows spectrum + background)
            selectionDomain = displayAndSelect(infoList, dataList, dataList_notSubtracted)

        elif request == "recalculate":
            selectionDomain = selectionDomains[i]
        else:
            print("warning, not a valid request")

        # Get total counts per second of the selection
        selectionCPS = getSelectionCPS(infoList, dataList, selectionDomain)

        # Get the measured activity of the source by dividing by the chance this emission occurs
        measuredActivity = 100 * selectionCPS / (emissionPeaks[i][2] ) # 100 because the given fraction is in percentages

        # Actual predicted activity of the source by correcting for the efficiency of the detector
        activity = measuredActivity / detectorEfficiency
        activityPerIsotope.append((emissionPeaks[i], selectionDomain, activity))
    return activityPerIsotope


def activityListToString(IsotopeActivity):
    isotopeInfo, selectionDomain, activity = IsotopeActivity
    return (f"{isotopeInfo[0]} at emission of {isotopeInfo[1]} keV ({isotopeInfo[2]}%)."
            f"\nSelection from channel {energyToChannel(infoList, selectionDomain[0])} ({selectionDomain[0]} keV) to channel {energyToChannel(infoList, selectionDomain[1])} ({selectionDomain[1]} keV)."
            f"\nActivity = {activity} cps.\n\n")


## Hardcoded constants

# Efficiency of the detector, constant (since we do not know the exact values)
detectorEfficiency = 0.10

sourceType = "Autunite" # Contains natural Uranium
if sourceType == "Autunite":
    sourceFile = "Spectra/AutuniteSpectrum.spe"

    # Multiple emission per isotope to increase accuracy. Source: https://nds.iaea.org/records/nfcbt-q6e23, page 36+
    emissionPeaks = [("Th-234", 63.3, 3.7), ("Th-234", 92.6, 5.2), ("Th-234", "63-93", 8.9), # 112.8, 0.24% (not easily visible)
                     ("U-235", 143.76, 11.0), ("U-235", 185.72, 57.2), # last with Ra-226 (186.2), decay of U-238, action?.
                     ("Pb-214", 242.0, 7.2), ("Pb-214", 295.22, 18.3), ("Pb-214", 351.93, 35.3),
                     ("Bi-214", 609.3, 45.2)] # [(isotope_name (str), energy_of_emission_peak (str/double), fraction_of_activity_in_percentages (double)), (...), ...]

    # Erases the peak when not yet added
    emissionPeaks = [peak for peak in emissionPeaks if peak[2] is not None]

else:
    print("Warning! Other elements need to be hardcoded in.")
    exit()


## Execution

# Get the spectrum
infoList, dataList_notSubtracted = getSpectrum(sourceFile)

# Subtract the background radiation
dataList = subtractBackground(infoList, dataList_notSubtracted)

request = "recalculate" # "new selection" (default) || "recalculate"
activityPerIsotope = getActivityPerIsotope(request, infoList, dataList, dataList_notSubtracted, emissionPeaks, detectorEfficiency)


# output
response = input("File selections.txt wil be overridden. Do you want to proceed?\n")
if response.lower().startswith("y"):
    filePath = "selections2.txt"

    print(f"Writing to {filePath} ...")

    # Output to text file
    file = open(filePath, "w")
    for IsotopeActivity in activityPerIsotope:
        file.write(activityListToString(IsotopeActivity))
else:
    print("File selections.txt wil not be overridden.\n")

    # Output to terminal
    for IsotopeActivity in activityPerIsotope:
        print(activityListToString(IsotopeActivity))
