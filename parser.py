# Functionality: Loop for each file and parse the data we want and append to same excel sheet
import pandas as pd
import json 
import xlwings as xw
import sys

tryPathCounter = 0
margin = -1000 
sign = 1
newMinValue = 1
previous_value = None
current_value = None

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

print("\n***Enter 0 to quit whenever prompted***")

desiredLocation = input("\nEnter named position with space between words:  ")

desiredLocation = desiredLocation.title() 

desiredLocation = desiredLocation.replace(" ","")

while desiredLocation not in offset:
    print("\n***Invalid location***")
    desiredLocation = input("\nEnter named position with space between words: ")
    desiredLocation = desiredLocation.title() 

    desiredLocation = desiredLocation.replace(" ","")

while True:
    targetFile = input("\nPath to JSON file: ")

    if (targetFile == '0'):
        print("Exiting...")
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

    desiredData = desiredData * 1000

    # Create dictionary that will be sent to excel
    desiredDataFrame = {
        'Module': [module],
        desiredLocation: [desiredData],
        'Offset' : [offset[desiredLocation]],
        'Margin' : [margin],
        'Sign' : [sign],
        'Datum' : [(desiredData + margin) * sign]
    }

    # Create Data Frame from dictionary, index -> ensures key-value parsed as row 
    frame = pd.DataFrame.from_dict(desiredDataFrame, orient='columns')

    # Open existing excel sheet
    try: 
        existingFrame = pd.read_excel('output.xlsx')
    except: 
        print("File does not exist.")
        continue

    # Can't write to open sheet, so this opens and closes if present 
    try: 
        excel = xw.Book('output.xlsx')
        excel.close()
    except:
        print("File does not exist.")

    current_value = frame['Datum'].min()

    if (previous_value is not None):
        if (previous_value > current_value):
            newMinValue = current_value
            frame.loc[0, 'Datum Min'] = [newMinValue]

    previous_value = current_value

# Append new frame to existing frame
    if (not existingFrame.empty):
        df_combined = existingFrame._append(frame, ignore_index = True)
        df_combined.to_excel('output.xlsx', index=False)
    else:
        frame.to_excel('output.xlsx', index=False)
    
    


