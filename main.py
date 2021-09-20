import os
from datetime import datetime
from typing import List

from openpyxl import Workbook, load_workbook #lib for managing excel https://openpyxl.readthedocs.io/en/stable/index.html 
from openpyxl.utils import get_column_letter
import pandas as pd
from pandas.core.dtypes.missing import notnull

cwd = os.getcwd()

NRIC_COLUMN = "Last 4 Alphanumeric of NRIC"
BANK_ACCOUNT_COLUMN = "Full Name as per NRIC/ Passport" # change to "Please enter Bank Account Holder's name" afterwards

def bank_GetData():
    
    ## open the workbook, create if doesn't exist
    BANK_PATH = cwd+"/bankStatements.xlsx"
    if not os.path.isfile(BANK_PATH): 
        wb=Workbook()
        ws1 = wb.active
        ws1.title = "Raw"
        wb.save(filename = BANK_PATH)
        return 0
    wb_panda = pd.read_excel(BANK_PATH,sheet_name="Raw",engine="openpyxl",dtype=str)
    # print("bank:\n")
    # print(wb_panda.head())

    #The bank statement always has `Total Debit Count` at the end of the document, so you can scan until there
    return wb_panda

def gf_GetData():
    """Get the Google Form Responses. It will be multiple column, with the key based on the fields of the form. Make this dynamic? Since google forms produces the first row as a fixed title name anyway"""
    GF_PATH = cwd+"/form2Responses.xlsx"

    # if the workbook doesnt exist create it. Ensure formatting is correct as well
    wb=Workbook()
    if not os.path.isfile(GF_PATH): 
        ws1 = wb.active
        ws1.title = "Raw"
        wb.create_sheet(title="Paid")
        wb.save(filename = GF_PATH)
        print("Please input the data")
        return 0
    # wb.active.title="Raw"
    # wb.save(filename=GF_PATH)
    
    # try again with pandas
    wb_panda = pd.read_excel(GF_PATH,sheet_name="Raw",engine="openpyxl", dtype=str)
    wb_panda = wb_panda.fillna('') # incase i nid to replace the nan with a blank

    # print("gf:\n")
    # print(wb_panda.head())

    return wb_panda

# find the different formats of payment. PayNow has a format, Paylah, BankTransfer. If majority of the cases are paynow/paylah, then why not i automate that first. To do a bank transfer the patient must have emailed the clinic.
def bank_Searcher():
    """detects the format of a paynow/paylah and puts the relevant fields in a new column"""

    return
            
def masterFilter(GF,Bank):
    """
    The core filtering logic.
    Compare every entry in the gf with the bank statements. Store in a list of possible matches, if more than one remains, unlock the next layer of search. 
    Do in this order..
    1. Check Full IC
    2. Check first name segment from bank account
    3. If multiple, check next name segment among survivors
    4. Still multiple, check if GF has multiple same name
    """

    # set output file
    output_path = outputFileName()
    matches =[]
    bank_transaction=[]

    # lists to hold the filtered rows
    paid_too_many_times=[]
    paid_for_others=[]
    no_payment_found=[]
    paid_correctly=[]
    paid_wrong_amount=[]


    for a in GF.index:
        gf_nric = GF[NRIC_COLUMN][a]
        name_list = GF[BANK_ACCOUNT_COLUMN][a]
        
        if search_NRIC(gf_nric,Bank):
            print(str(GF[BANK_ACCOUNT_COLUMN][a])+" found by nric")
            bank_duplicate, bt = checkAgainstBank(name_list,Bank)
            bank_transaction.append(bt)
            if bank_duplicate:
                print(str(GF[BANK_ACCOUNT_COLUMN][a])+"Paid too many times")
                paid_too_many_times.append(a)
            else:
                matches.append(a)
        elif search_Name(name_list,Bank):
            print(str(GF[BANK_ACCOUNT_COLUMN][a])+" found by name")
            bank_duplicate, bt = checkAgainstBank(GF[BANK_ACCOUNT_COLUMN][a],Bank)
            bank_transaction.append(bt)
            if bank_duplicate:
                if checkAgainstGF(GF,a):
                    print(str(GF[BANK_ACCOUNT_COLUMN][a])+"Paid for others")
                    paid_for_others.append(a)
                else:
                    print(str(GF[BANK_ACCOUNT_COLUMN][a])+"Paid too many times")
                    paid_too_many_times.append(a)
            else:
                matches.append(a)
        else: 
            print(str(GF[BANK_ACCOUNT_COLUMN][a])+"No payment found")
            no_payment_found.append(a)
        a+=1
    print("matches: "+str(matches))
    ok=0
    for b in matches:
        for c in range(6):
            try: (Bank.iloc[bank_transaction[b]+c,0])
            except: 
                print("continue")
                break
            if "SGD 98" in Bank.iloc[bank_transaction[b]+c,0]:
                    print(str(GF[BANK_ACCOUNT_COLUMN][b])+" OK")
                    paid_correctly.append(b)
                    ok = 1
                    break
        if ok==0:
            print(str(GF[BANK_ACCOUNT_COLUMN][b])+" paid wrong amount")
            paid_wrong_amount.append(b)
            
            # else: print(str(GF[BANK_ACCOUNT_COLUMN][b])+"Paid wrong amount")
            # else: print(str(GF[BANK_ACCOUNT_COLUMN][b])+"Technical Error: can't find amount paid")
    
    return paid_too_many_times, paid_for_others, no_payment_found, paid_correctly, paid_wrong_amount



def search_NRIC(gf_nric, Bank):
    
    charnum=len(gf_nric)
    if charnum>0:
        for bah in Bank.index:
            if gf_nric.lower() in Bank.iloc[bah,0].lower():
                return 1
        else:
            return 0


def search_Name(name_list, Bank):
    
    for bah in Bank.index: 
        if name_list.lower() in Bank.iloc[bah,0].lower():
            return 1
    else:
        return 0

def checkAgainstBank(name_list,Bank):
    count=0
    address=999

    for bah in Bank.index:
        if name_list.lower() in Bank.iloc[bah,0].lower():
            count+=1
            if count==1:
                address=bah
    if count>1:
        return 1
    else:
        return 0,address

def checkAgainstGF(GF,r):
    q=0
    count=0

    print(GF[BANK_ACCOUNT_COLUMN][r].lower())
    for q in GF.index:
        if GF[BANK_ACCOUNT_COLUMN][r].lower() in GF[BANK_ACCOUNT_COLUMN][q].lower():
            count+=1
        q+=1
    if count>1:
        return 1
    else:
        return 0
            


def outputFileName():
    """Before the start of your printing, create a new output file that doesn't overwrite the existing (even those made in the same day)"""
    # check if the default name exists
        #if so, then iterate on the increments until you find one that's free. Set the output as that
    now = datetime.now().strftime("%Y%m%d")
    x=0
    output_file = now+"_output_"+str(x)
    OUTPUT_PATH = cwd+"/"+output_file+".xlsx"
    while 1:
        if os.path.isfile(OUTPUT_PATH):
            x+=1
            output_file = now+"_output_"+str(x)
            OUTPUT_PATH = cwd+"/"+output_file+".xlsx"
        else:
            break
    wb=Workbook()
    ws1 = wb.active
    ws1.title = "Paid Correctly"
    wb.create_sheet(title="Paid too many times")
    wb.create_sheet(title="Paid for others")
    wb.create_sheet(title="No payment found")
    wb.create_sheet(title="Paid wrong amount")
    wb.save(filename=OUTPUT_PATH)
    wb.close()

    return OUTPUT_PATH

def printMatches(GF,data,sheet,output_path):

    mylist =[]

    for i in data:
        mylist.append(GF.loc[i])

    df = pd.DataFrame(mylist,columns=GF.columns)
    

    # df = pd.DataFrame(mylist)
    # print(df.head())

    # # if the workbook doesnt exist create it. Ensure formatting is correct as well
    wb=load_workbook(output_path)
    # ws = wb[sheet]
    writer=pd.ExcelWriter(output_path, engine='openpyxl')
    writer.book = wb

    writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
    df.to_excel(writer,sheet)
    writer.save()
    



if __name__ == "__main__":
    # ## lets deal with dataframe because we dunnid to do arithmetic, but we wna edit them alot.

    # # #also read the fields in the googleform. Search the first row for headers.
    # myGF = gf_GetData()
    # # first i need to get all the rows in bankStatements.xlsx and put it in a big list
    # myBank = bank_GetData()

    # #identify a paynow/paylah transaction in the long array, and pull the NRIC (if available) and each segment of their name.
    

    
    # masterFilter(myGF,myBank)

    # ## if at the end of this you still have multiple names, this means multiple payment. Put them in another sheet called "Multiple Payment"

    # ## the rest of them put it in the "Not Paid" sheet

    ###########################
    ##   Space for Testing   ##
    ###########################
    x = gf_GetData()
    y = bank_GetData()
    paid_too_many_times, paid_for_others, no_payment_found, paid_correctly, paid_wrong_amount = masterFilter(x,y)
    # search_Name(x,y,search_NRIC(x,y))
   
    mypath = outputFileName()
    print(mypath)

    printMatches(x,paid_correctly,"Paid Correctly",mypath)
    printMatches(x,paid_too_many_times,"Paid too many times",mypath)
    printMatches(x,paid_for_others,"Paid for others",mypath)
    printMatches(x,no_payment_found,"No payment found",mypath)
    printMatches(x,paid_wrong_amount,"Paid wrong amount",mypath)

