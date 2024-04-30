# Functionality: Loop for each file and parse the data we want and append to same excel sheet
import pandas as pd
import json 
import xlwings as xw
import os

count = 0
margin = -1000 
sign = 1
output_file = 'output.xlsx'
flag = None

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

def get_desired_location():
    global desired_location 
    desired_location = input("\nEnter named position with space between words:  ")

    # Capitalize first letter of each word
    desired_location = desired_location.title() 

    # Cut out space in between words
    desired_location = desired_location.replace(" ","")

get_desired_location()

# Open existing excel sheet otherwise create empty frame 
try: 
    existing_frame = pd.read_excel(output_file, sheet_name=[procedures[desired_location]])
    existing_frame = pd.DataFrame.from_dict(existing_frame, orient='columns')
except: 
    existing_frame = pd.DataFrame() 

# Can't write to open sheet, so this opens and closes if present 
try: 
    excel = xw.Book(output_file)
    excel.close()
except:
    pass

while desired_location not in offset:
    print("\n***Invalid location***")
    get_desired_location()

while True:
    target_file_path = input("\nPath to JSON file: ")

    if (target_file_path == 0):
        print("\nExiting...")
        break

    # Load JSON as a dictionary 
    while True:
        try: 
            with open(target_file_path, 'r') as file:
                data = json.load(file) 
                break
        except FileNotFoundError:
                count += 1
                if (count == 2):
                    print("\nExiting...")
                    flag = True
                    break
                else:
                    print("\n***Invalid path***")
                
        target_file_path = input("\nPath to JSON file: ")

        if ("\\" not in target_file_path):
            print("\nExiting...")
            flag = True
            break
    
    if flag: 
        break

    # Parse module from path
    target_split = target_file_path.split("\\")
    module = target_split[6]

    if "hq" in module:
        margin = 1000

    # Path to data we want
    desired_data = data['TubeRobotZAxis']['NamedPositions'][desired_location]
    desired_data = round(desired_data * 1000)

    # Add values to list to display to excel
    moduleList.append(module)
    positionValueList.append(desired_data)
    datumList.append((desired_data + margin) * sign)
    marginList.append(margin)
    signList.append(sign)
    offsetList.append(offset[desired_location])
    emptyList.append(None)

# Create dictionary that will be sent to excel
desiredDataFrame = {
    'Module': moduleList,
    desired_location: positionValueList,
    'Offset' : offsetList,
    'Margin' : marginList,
    'Sign' : signList,
    'Datum' : datumList,
    '' : emptyList
}

# Create Data Frame from dictionary, index -> ensures key-value parsed as row 
frame = pd.DataFrame.from_dict(desiredDataFrame, orient='columns')

# Calculation for calcFrame
min_value = frame['Datum'].min()
datum_sum = frame['Datum'].sum()
num_datums = frame['Datum'].count()
avg_datum = round(datum_sum / num_datums)
datum_delta = avg_datum - min_value


# Create Data Frame 
calcFrame = pd.DataFrame(
    {
        'Datum Avg': [avg_datum],
        'Datum Min': [min_value],
        'Datum Delta': [datum_delta]
    }
)

#Create Excel Writer 
if os.path.exists(output_file):
    with pd.ExcelWriter(output_file,engine='openpyxl', mode='a', if_sheet_exists='replace') as writer: 
        df_combined = existing_frame._append(frame, ignore_index = True)
        final_frame = pd.concat([df_combined, calcFrame], axis=1)
        final_frame.to_excel(writer, sheet_name=procedures[desired_location], index=False)
        
        total_sheets = len(writer.book.sheetnames)

        for i in range(total_sheets):
            workbook = writer.book 
            worksheet = workbook.worksheets[i]
            worksheet.column_dimensions['b'].width = len(desired_location) + 2
            worksheet.column_dimensions['H'].width = 12
            worksheet.column_dimensions['I'].width = 12
            worksheet.column_dimensions['J'].width = 12
else: 
    with pd.ExcelWriter(output_file,engine='openpyxl') as writer: 
            final_frame = pd.concat([frame, calcFrame], axis=1)
            final_frame.to_excel(writer, sheet_name=procedures[desired_location], index=False)  

            workbook = writer.book 
            worksheet = workbook.worksheets[0]
            worksheet.column_dimensions['b'].width = len(desired_location) + 2 
            worksheet.column_dimensions['H'].width = 12
            worksheet.column_dimensions['I'].width = 12
            worksheet.column_dimensions['J'].width = 12  

