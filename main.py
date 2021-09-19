import os
from datetime import datetime

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

    # compare the NRIC value
    nric_matches,nric_multi_payments = search_NRIC(GF, Bank)
    print("matches: "+str(nric_matches))
    print("duplicate: "+str(nric_multi_payments))

    # printMatches(nric_matches,output_path)

    if bool(nric_multi_payments) != 0:
        duplicates = list_duplicates(GF,Bank,nric_multi_payments,NRIC_COLUMN)

    # search the list again for remainder, or everthing if no match 
    name_matches, name_multi_payments = search_Name(GF, Bank, nric_matches)    

    all_matches = nric_matches.extend(name_matches)

    # printMatches(name_matches,output_path)
    # check that there are no multiple match
    if bool(name_multi_payments) != 0:
        printMatches(name_matches,output_path)
    # else:
        # check the gf itself if there are multiple entries with this name. This means the guy paid for other ppl 
    
    printMatches(all_matches,output_path)


def search_NRIC(GF, Bank):
    
    # holder list for matched records
    matches=[]
    multi_payments=[]
    search_is_definitive=1

    # find the column containing NRIC
    # for x in GF.columns.values:
    #     if NRIC_COLUMN in x: # make sure it's the NRIC 
    #            NRIC_Column = x

    for j in GF.index: # iterate throught the rows

        charnum = len(GF[NRIC_COLUMN][j]) # i kno how long they enter liao, incase clare dun check for length agn
            
        if charnum>0: 
            gf_nric = GF[NRIC_COLUMN][j]
            
            for bah in Bank.index:
                # if Bank.iloc[bah,0].find(gf_nric) != -1:
                if gf_nric.lower() in Bank.iloc[bah,0].lower():
                    # print(Bank.iloc[bah,0])
                    # print(gf_nric)                     
                    if j in matches:
                        try: 
                            multi_payments.append(j)
                        except: 
                            multi_payments.insert(j,0)
                    try: 
                        matches.append(j)
                    except: 
                        matches.insert(j,0)
                    # print("main: "+str(Bank.iloc[bah,0]))
                    

    # remove duplicates from matches
    holder=[]
    i=0
    for key in matches:
        if matches[i] not in multi_payments:
            try: 
                holder.append(key)
            except: 
                holder.insert(key,0)
        i+=1
    matches=holder

                
    # if there are no matches, return the entire search list
    if len(matches) == 0:
        # matches = GF.index ## I shouldn't do this, it outputs: "RangeIndex(start=0, stop=10, step=1)"
        for q in GF.index:
            print(q)
            try: matches.append(q)
            except: matches.insert(q,0)

    return matches, multi_payments


def search_Name(GF, Bank, nric_matches):
    matches = []
    multi_payments=[]
    same_name = []
    
    for j in GF.index:

        name_list=GF[BANK_ACCOUNT_COLUMN][j] # output should be a list of name segments

        
        # For every line in gf, search entire bank statement for name.
        for bah in Bank.index:
            if name_list.lower() in Bank.iloc[bah,0].lower():
                # print(Bank.iloc[bah,0])
                if j in nric_matches: # check the list you alr have if you've matched before
                    try: multi_payments.append(j)
                    except: multi_payments.insert(j,0)
                try: matches.append(j)
                except: matches.insert(j,0)
        
    # remove duplicates from matches
    holder=[]
    i=0
    for key in matches:
        if matches[i] not in multi_payments:
            try: 
                holder.append(key)
            except: 
                holder.insert(key,0)
        i+=1
    matches=holder
    
    # if there are no matches, return the entire searched list
    if len(matches) == 0:
        matches = nric_matches

    # print(matches)
    return matches, multi_payments

# THE SPLIT NAME LOGIC IS FLAWED!
# def gf_SplitName(A):
#     """split the name into strings, separated by spaces and commas"""
#     count=0
#     name_components_bucket=[]
#     names=["john","tan"] #these are the names of the patients
#     for i in A:
#         if i == " " or i == ",":
#             names.append(char_bucket)
#         char_bucket.append(i)

def list_duplicates(GF,Bank,nric_duplicates,column):
    matches=[[],[]]
    for i in nric_duplicates:
        gr_nric = GF[column][i]

        for bah in Bank.index:
            if gr_nric.lower() in Bank.iloc[bah,0].lower():
                try: 
                    matches[0].append(i)
                    matches[1].append(bah)
                except: 
                    matches[0].insert(i,0)
                    matches[1].insert(bah,0)

    return matches

def outputFileName():
    """Before the start of your printing, create a new output file that doesn't overwrite the existing (even those made in the same day)"""
    # check if the default name exists
        #if so, then iterate on the increments until you find one that's free. Set the output as that
    now = datetime.now().strftime("%Y%m%d")
    output_file = now+"output"
    OUTPUT_PATH = cwd+"/"+output_file+".xlsx"

    return OUTPUT_PATH

def printMatches(data,output_path):

    # if the workbook doesnt exist create it. Ensure formatting is correct as well
    output_path
    wb=Workbook()
    return


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
    masterFilter(x,y)
    # search_Name(x,y,search_NRIC(x,y))

