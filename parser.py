# Functionality: Loop for each file and parse the data we want and append to same excel sheet
import pandas as pd
import json 
import xlwings as xw
import sys
import os

count = 0
tryPathCounter = 0
margin = -1000 
sign = 1
minValue = 0
datumTotal = 0
datumAvg = 0
outputFile = 'output.xlsx'

# Lists for Data Frame
moduleList = []
positionValueList = []
datumList = []
marginList = []
signList = []
offsetList = []
emptyList = []

offset = {
    "CarrierRackHeight" : -12500,
    "ConveyorHeight" : -12500, 
    "DryStationHeight" : -1000, 
    "FlipperHeight" : -20000,
    "EscalatorTubeHeight" : 2000,
    "OpenTubeStationHeight" : -12500,
    "ResuspensionTubeHeight" : 1000, 
    "StainBath" : -21000,
    "TubeRackHeight" : -12500
}

procedures = {
    "CarrierRackHeight": "Racks",
    "ConveyorHeight": "Conveyor",
    "DryStationHeight" :"Dry Station", 
    "FlipperHeight" : "Flipper",
    "EscalatorTubeHeight" : "Escalator",
    "OpenTubeStationHeight" : "Open Tube Station",
    "ResuspensionTubeHeight" : "Resuspension", 
    "StainBath" : "StainBath",
    "TubeRackHeight" : "Tubes"
}

print("\n***Enter 0 to quit whenever prompted***")

desiredLocation = input("\nEnter named position with space between words:  ")

desiredLocation = desiredLocation.title() 

# Cut out space in between words
desiredLocation = desiredLocation.replace(" ","")

# Open existing excel sheet otherwise create empty frame 
try: 
    existingFrame = pd.read_excel(outputFile, sheet_name=[procedures[desiredLocation]])
except: 
    existingFrame = pd.DataFrame() 

# Can't write to open sheet, so this opens and closes if present 
try: 
    excel = xw.Book(outputFile)
    totalSheets = len(excel.sheet_names)
    excel.close()
except:
    pass

while desiredLocation not in offset:
    print("\n***Invalid location***")
    desiredLocation = input("\nEnter named position with space between words: ")
    desiredLocation = desiredLocation.title() 
    desiredLocation = desiredLocation.replace(" ","")

while True:
    targetFile = input("\nPath to JSON file: ")

    if (targetFile == '0'):
        print("\nExiting...")
        break

    # Load JSON as a dictionary 
    while True:
        try: 
            with open(targetFile, 'r') as file:
                data = json.load(file) 
                break
        except FileNotFoundError:
                tryPathCounter += 1
                if (tryPathCounter == 2):
                    print("\nExiting...")
                    sys.exit()
                else:
                    print("\n***Specified path does not exist***")
                    print("\n**One more chance, so make it right**")
                
        targetFile = input("\nPath to JSON file: ")

        if (targetFile == 0):
            break
        
    # Parse module from path
    splitTargetFile = targetFile.split("\\")
 
    module = splitTargetFile[6]

    # Path to data we want
    desiredData = data['TubeRobotZAxis']['NamedPositions'][desiredLocation] 

    desiredData = round(desiredData * 1000)

    moduleList.append(module)
    positionValueList.append(desiredData)
    datumList.append((desiredData + margin) * sign)
    marginList.append(margin)
    signList.append(sign)
    offsetList.append(offset[desiredLocation])
    emptyList.append(None)

    # Create dictionary that will be sent to excel
    desiredDataFrame = {
        'Module': moduleList,
        desiredLocation: positionValueList,
        'Offset' : offsetList,
        'Margin' : marginList,
        'Sign' : signList,
        'Datum' : datumList,
        '' : emptyList
    }

    # Create Data Frame from dictionary, index -> ensures key-value parsed as row 
    frame = pd.DataFrame.from_dict(desiredDataFrame, orient='columns')

    # Compare and store minimum value
    minValue = frame['Datum'].min()

    # Get total of all datums to calculate average 
    datumTotal += sum(datumList)
    count += 1

    datumAvg = round((datumTotal / count))

# Create Data Frame 
calcFrame = pd.DataFrame(
    {
        'Datum Avg': [datumAvg],
        'Datum Min': [minValue],
        'Datum Delta': [datumAvg - minValue]
    }
)

#Create Excel Writer 
if os.path.exists(outputFile):
    with pd.ExcelWriter(outputFile,engine='openpyxl', mode='a') as writer: 
        # Append new frame to existing frame
        df_combined =existingFrame._append(frame, ignore_index = True)
        finalFrame = pd.concat([df_combined, calcFrame], axis=1)
        finalFrame.to_excel(writer, sheet_name=procedures[desiredLocation], index=False)
        
        for i in range(totalSheets + 1):
            workbook = writer.book 
            worksheet = workbook.worksheets[i]
            worksheet.column_dimensions['b'].width = len(desiredLocation) + 2
            worksheet.column_dimensions['H'].width = 12
            worksheet.column_dimensions['I'].width = 12
            worksheet.column_dimensions['J'].width = 12
else: 
    with pd.ExcelWriter(outputFile,engine='openpyxl') as writer: 
            # If the file doesn't exist 
            finalFrame = pd.concat([frame, calcFrame], axis=1)
            finalFrame.to_excel(writer, sheet_name=procedures[desiredLocation], index=False)  

            workbook = writer.book 
            worksheet = workbook.worksheets[0]
            worksheet.column_dimensions['b'].width = len(desiredLocation) + 2 
            worksheet.column_dimensions['H'].width = 12
            worksheet.column_dimensions['I'].width = 12
            worksheet.column_dimensions['J'].width = 12  

