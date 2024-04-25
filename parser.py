import pandas as pd
import json 
import xlwings as xw
import sys

exit = False
row = 0
while (exit != True):
    targetFile = input("\nPath to JSON file: ")

    # Load JSON as a dictionary 
    try: 
        with open(targetFile, 'r') as file:
            data = json.load(file) 
    except FileNotFoundError:
            print("\n***Specified path does not exist***")
            sys.exit()

    desiredData = data['TubeRobotZAxis']['NamedPositions']['CarrierRackHeight'] # Path to data we want

    # Create dictionary that will be sent to excel
    desiredDataFrame = {
        'Value': desiredData
    }
    print(desiredDataFrame)

    # Create Data Frame from dictionary, index -> ensures key-value parsed as row 
    frame = pd.DataFrame.from_dict(desiredDataFrame, orient='index')

    # Open existing excel sheet
    existingFrame = pd.read_excel('output.xlsx')

    # Needed this in order to overwrite excel sheet since it can't be done while open. So open it and close it. 
    try: 
        excel = xw.Book('output.xlsx')
        excel.close()
    except Exception as e:
        print (e)

    # Insert data into specified row and column 
    existingFrame.iloc[row,1] = desiredData
    
    # Since we are entering multiple entries, move on to next row
    row += 1

    # Append new frame to existing frame 
    df_combined = existingFrame._append(frame, ignore_index = True)

    # Output to excel sheet
    df_combined.to_excel('output.xlsx', index=False)
