import subprocess       
import pandas as pd     
from time import sleep as wait            
import os               
import sys
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter import messagebox
import functions as func
import pyautogui as ag
import logging
from datetime import datetime



# Set Path variables
projectFolder = r"C:\Projetos_Lighthouse\STO_Automation"
mainFolder = projectFolder + r"\main"
inputFolder = projectFolder + r"\input"
dataFolder = projectFolder + r"\data"
archiveFolder = dataFolder + r"archive"
scriptsFolder = dataFolder + r"\scripts"
tempFolder = dataFolder + r"\temp"
userdataFolder = projectFolder + r"\userdata"
logFolder = dataFolder + r"\log"
print(logFolder)

# Configure logging
func.setup_logging(logFolder)

print('Process started')

print("DEBUG===============================> Variables assigned successfully\n")

# Close SAP
os.system("taskkill /f /im saplogon.exe >nul 2>&1")

# delete files from temp folder
func.delete_files(tempFolder)
print("Files on temp folder were deleted\n")

# File with user credentials to login into SAP
userdataPath = userdataFolder + r"\Credentials.xlsx"
# Read userdata
df = pd.read_excel(userdataPath, header=None)
login = df.iloc[1, 0]
senha = df.iloc[1, 1]

print("DEBUG===============================> User data successfully read\n")


# Step1: Read input file:
# Prompt to user select the file that will be processed
inputFilePath = func.select_file()

if inputFilePath:
    print("The file " + inputFilePath + " was selected\n")
else:
    print("No file was selected")
    messagebox.showwarning("Error", "No file was selected. Please select a file to continue.")
    os._exit(0)

print("DEBUG===============================> Input file was successfully selected.\n")

# Step2: Separate each individual order into unique files:

# check if the message column exists and creates it if not
func.ensure_message_column(inputFilePath)

# load file data
df = pd.read_excel(inputFilePath, engine='openpyxl')


# get all unique values from ID column
unique_ids = df['Unique ID'].unique()


# Create new file on temp folder for each unique ID
for unique_id in unique_ids:
    df_filtered = df[df['Unique ID'] == unique_id]
    output_file = tempFolder + f'\delivery_{unique_id}.xlsx'
    df_filtered.to_excel(output_file, index=False, engine='openpyxl')

print("DEBUG===============================> Files were separated successfully\n")

#Login SAP
sap_path = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\SAP Front End\SAP Logon.lnk"
subprocess.Popen(['start', '', sap_path], shell=True)
wait(4)
func.RunVB(scriptsFolder + r"\loginSAP.vbs", ["9A. Symphony ECC Production", login, senha])

print("DEBUG===============================> Logged in on SAP\n")

# Step3: Loop into individual files
for filename in os.listdir(tempFolder):
    file_fullpath = os.path.join(tempFolder, filename)
    print("Processing file: "+ file_fullpath)

    fileData = pd.read_excel(file_fullpath, engine='openpyxl')
    print(fileData)

    wait(3)

    # Access ME21N transaction
    func.RunVB(scriptsFolder + r"\openME21N.vbs",[])

    wait(1)

    func.RunVB(scriptsFolder + r"\openHeaderSection.vbs",[])

    # Get "Origin" and "Ship to" data
    first_row = fileData.iloc[0]
    Origin = first_row['Origin']
    Ship_To = first_row['Ship To']
    Due_Date = first_row['Due Date']
    formatDate = func.formatDate(Due_Date)
    orderMessage = first_row['Message']

    findString = "Intra-Company STPO created under the number"
    findError = "ERROR"
    if findString in str(orderMessage) or findError in str(orderMessage):
        print("Order already processed, proceeding to the next...\n")
        # Order already processed
        continue

    # fill header fields
    func.RunVB(scriptsFolder + r"\headerFields.vbs", [])
        
    # supllying plant
    func.RunVB(scriptsFolder + r"\supPlant.vbs", [Origin])
    # formatDate = "06/30/2025"
    # update date values
    func.RunVB(scriptsFolder + r"\dateAndShip.vbs", [formatDate, Ship_To])

    
    # Loop on table items
    for index, row in fileData.iterrows():
        Unique_ID = row['Unique ID']
        Material = row['Material']
        qty = row['Total Confirmed Cases']
        Due_Date = row['Due Date']
        Batch = row['Batch']

        #format material number to ensure it have at least 5 digits
        Material = func.format_material(Material)


        
        print(f"-------------Processing line {index + 1}:-----------\n")
        print(f"  Unique ID: {Unique_ID}\n")
        print(f"  Origin: {Origin}\n")
        print(f"  Material: {Material}\n")
        print(f"  Total Confirmed Cases: {qty}\n")
        print(f"  Ship To: {Ship_To}\n")
        print(f"  Due Date: {Due_Date}\n")
        print(f"  Batch: {Batch}\n")

        # Fill table fields
        func.RunVB(scriptsFolder + r"\tableFields.vbs", [index, Material, qty])

        # Fill batch field if there is an available batch code
        print("batch: " + str(Batch) + "\n")
        outbatch = func.RunVB(scriptsFolder + r"\fillBatch.vbs", [index, Batch])
        print("outBatch = " + outbatch + "\n")



    print("DEBUG===============================> Out of the items loop\n")
    # Out of the items loop
    ag.press("ENTER")
    
    wait(2)

    # sorry for the poor quality of this validation part - basically it will see if a message appeared on the lower bar of the SAP screen
    validMsg2 = func.RunVB(scriptsFolder + r"\getGenNumber.vbs", []) # Get gen number script gets the text written on the SAP lower bar, from the error messages to the generated numbers                                               
    if validMsg2:
        # error on item. 
            print(f'Message appeared on SAP lower bar: {validMsg2}\n')
            errMsgWhenSave = f"ERROR: {validMsg2}" # save this because it will be the reason of the error that will be thrown when trying to save the order
    

    validMsg = func.RunVB(scriptsFolder + r"\validMessage.vbs", []) # Validate if "wnd[1]/usr" exists
    if validMsg == "True":
        # some message appeared when selecting the item
        # close message:
        func.RunVB(scriptsFolder + r"\closeWin.vbs", [1])
        # See if there is an error msg
        validErrMsg = func.RunVB(scriptsFolder + r"\getGenNumber.vbs", [])
        if validMsg:
            # error on item. stop order insertion
            print('An error occurred while trying to process this order\n')
            func.RunVB(scriptsFolder + r"\cancelOrder.vbs", [])
            func.update_excel(inputFilePath, Unique_ID, validErrMsg)
            continue

    wait(3)

    # Click the Save button
    func.RunVB(scriptsFolder + r"\saveOrder.vbs", [])

    wait(2)

    validMsg = func.RunVB(scriptsFolder + r"\validMessage.vbs", []) # Validate if "wnd[1]/usr" appeared after trying to save the order
    if validMsg == "True":
        # some message appeared when selecting the item
        # See if there is an error msg
        wait(1)
        validErrMsg = func.RunVB(scriptsFolder + r"\getGenNumber.vbs", [])# validate if there is a message on SAP lower bar
        if errMsgWhenSave:
            # Error window ("wnd[1]/usr") appeared and there was an error previously, while trying to add the items. Log the error from before
            print(f'An error occurred while trying to process this order: {errMsgWhenSave}\n')
            # close message:
            func.RunVB(scriptsFolder + r"\closeWin.vbs", [1])
            wait(1)
            func.RunVB(scriptsFolder + r"\cancelOrder.vbs", [])
            func.update_excel(inputFilePath, Unique_ID, errMsgWhenSave)
            continue
        elif validErrMsg:
            # no error before, but one appeared now. log the error that just appeared
            validErrMsg = f"ERROR: {validErrMsg}"
            func.RunVB(scriptsFolder + r"\closeWin.vbs", [1])
            wait(1)
            func.RunVB(scriptsFolder + r"\cancelOrder.vbs", [])
            func.update_excel(inputFilePath, Unique_ID, validErrMsg)

           
        
        
    # no errors ocurred, save the order
    # Validate if there is a confirmation message and click save if there is
    validWin = func.RunVB(scriptsFolder + r"\validSaveBtn.vbs", [])

    if validWin == "True":
        func.RunVB(scriptsFolder + r"\clickFinalBtn.vbs", [])

    wait(3)

    # Return the number of the generated order
    orderCreated = func.RunVB(scriptsFolder + r"\getGenNumber.vbs", [])

    if "Intra-Company STPO created under the number" in orderCreated:
        print("Generated Order Number: " + orderCreated + "\n")
        func.update_excel(inputFilePath, Unique_ID, orderCreated+ "\n")
        # success scenario
        
    else:
        # error
        print("An error occured when trying to save the order: " + orderCreated + "\n")
        func.update_excel(inputFilePath, Unique_ID, orderCreated)


    

    # func.RunVB(scriptsFolder + r"closeWin.vbs", [1])
    func.RunVB(scriptsFolder + r"\sapMenu.vbs", [])
    wait(5)
    func.RunVB(scriptsFolder + r"\sapMenu.vbs", [])

    validSaveBtn2 = func.RunVB(scriptsFolder + r"\validSaveBtn.vbs", [])
    print(validSaveBtn2)
    func.RunVB(scriptsFolder + r"\dontSave.vbs", [])

        


        
    print("DEBUG===============================> Order Created, proceeding to next file...\n")
    # sys.exit()

# Next loop...
print("Finished\n")