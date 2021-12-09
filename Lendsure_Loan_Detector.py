import tkinter
import pandas as pd 
from tkinter import Button, messagebox, ttk, filedialog  
from ttkthemes import ThemedTk
import os
import datetime
import logging
import sys
import openpyxl

##########   GUI DESIGN   ##########
root = ThemedTk(theme="breeze")
root.title('Lendsure Duplicate Loan Detector')
#root.geometry("400x150")
root.iconbitmap(r"icon path")

#   VARIABLES
masterLoanFile = r"master file path"
sourceData = pd.read_csv(masterLoanFile)
mspLoanNums = sourceData['Loan Number'].tolist()
now = str(datetime.datetime.now())
now = (now[:10]+"_"+now[11:16]).replace(':', '.')
logging.basicConfig(filename=r"log file path", level=logging.INFO)
username = os.getlogin()


#   FUNCTIONS

def browse_cmd():

    """Opens file explorer browse dialogue box for user to search for files in GUI."""
    root.filename = filedialog.askopenfilename()
    #root.filename = filedialog.askopenfilename(initialdir=user_folder, title='Select a File', filetypes=(("All files", "*.*"),("Excel files", "*.xlsx")))
    filePathEntry.insert(0, root.filename)
    return None



def loanSearch_cmd():
    file = filePathEntry.get()
    submitButton["state"] = "disabled"
    if file == '':
        tkinter.messagebox.showerror('Empty File Path',
        'Please enter a file path or select the browse button to find your data file.')
        submitButton["state"] = "enable"
        return
    
    finaldataFolder = '/'.join(file.split('/')[0:-1])
    
    lendsureFile = pd.read_excel(file, engine='openpyxl')
    
    if 'Rushmore Loan #' not in lendsureFile.columns:
        tkinter.messagebox.showerror('Incorrect File','Please select a Lendsure final data file.')
        filePathEntry.delete(0,'end')
        submitButton["state"] = "enable"
        return


    
    loanDict ={'Rushmore Loan #':[], 'Prior Loan Number':[]}
    
    try:
        for i, loan in enumerate(lendsureFile['Rushmore Loan #']):
            if loan in mspLoanNums:
                loanDict['Rushmore Loan #'].append(loan)
                loanDict['Prior Loan Number'].append(lendsureFile['Loan Number'][i])
    except Exception as e:
        logging.info("User {} received the following error on {}: {}".format(username, now, e))
        loanDict ={'Rushmore Loan #':'none', 'Prior Loan Number':'none'}

    duplicateLoans = pd.DataFrame(loanDict, columns=['Rushmore Loan #', 'Prior Loan Number'], index=None)

    try:
        duplicateLoanInFile = pd.concat(g for _, g in lendsureFile.groupby("Rushmore Loan #") if len(g) > 1).iloc[:, 0:6]
    except Exception as e:
        logging.info("User {} received the following error on {}: {}".format(username, now, e))
        emptyDict ={'Rushmore Loan #':[], 'Prior Loan Number':[]}
        duplicateLoanInFile = pd.DataFrame(emptyDict, columns=['Rushmore Loan #', 'Prior Loan Number'], index=None)


    finalFile = pd.ExcelWriter(finaldataFolder + "/duplicateLoans.xlsx", engine='xlsxwriter')

    duplicateLoans.to_excel(finalFile, index=False, sheet_name='MSP duplicate loans')
    duplicateLoanInFile.to_excel(finalFile, index=False, sheet_name='Duplicate loans within the file')

    finalFile.save()
    
    filePathEntry.delete(0,'end')
    submitButton["state"] = "enable"


submitButton = ttk.Button(root, text="Submit", command=loanSearch_cmd)
filepathLabel = ttk.Label(root, text="Please select the Lendsure final data file.")
filePathEntry = ttk.Entry(root,  width=50 )
browseButton = ttk.Button(root, text='Browse', command= browse_cmd)
filepathLabel.grid(row=0, column=0, pady=10, padx=10)
filePathEntry.grid(row=1, column=0, pady=10, padx=10)
browseButton.grid(row=1, column=1, pady=10, padx=10)
submitButton.grid(row=2, column=0, padx=10, pady=10)




root.mainloop()
