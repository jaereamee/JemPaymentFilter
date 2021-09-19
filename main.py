import os

from openpyxl import Workbook, load_workbook #lib for managing excel https://openpyxl.readthedocs.io/en/stable/index.html 
from openpyxl.utils import get_column_letter
import pandas as pd

cwd = os.getcwd()

def bank_GetData():
    
    ## open the workbook, create if doesn't exist
    BANK_PATH = cwd+"/bankStatements.xlsx"
    if not os.path.isfile(BANK_PATH): 
        wb=Workbook()
        ws1 = wb.active
        ws1.title = "Raw"
        wb.save(filename = BANK_PATH)

    
    
    
    #The bank statement always has `Total Debit Count` at the end of the document, so you can scan until there
    return

# i also need to read the info from google form
def gf_GetData():
    """it will be multiple column, with the key based on the fields of the form. Make this dynamic? Since google forms produces the first row as a fixed title name anyway"""
    GF_PATH = cwd+"/form2Responses.xlsx"

    # if the workbook doesnt exist create it. Ensure formatting is correct as well
    wb=Workbook()
    if not os.path.isfile(GF_PATH): 
        print("nth")
        ws1 = wb.active
        ws1.title = "Raw"
        wb.create_sheet(title="Paid")
        wb.save(filename = GF_PATH)
        print("Please input the data")
        return 0
    # wb.active.title="Raw"
    # wb.save(filename=GF_PATH)
    
    # try again with pandas
    wb_panda = pd.read_excel(GF_PATH,sheet_name="Raw",engine="openpyxl")
    # df_panda = wb_panda.parse("Raw")
    print(wb_panda.head())

    return wb_panda

# find the different formats of payment. PayNow has a format, Paylah, BankTransfer. If majority of the cases are paynow/paylah, then why not i automate that first. To do a bank transfer the patient must have emailed the clinic.
def bank_Searcher():
    """detects the format of a paynow/paylah and puts the relevant fields in a new column"""

    return

def gf_SplitName(A):
    """split the name into strings, separated by spaces and commas"""
    count=0
    name_components_bucket=[]
    names=["john","tan"] #these are the names of the patients
    for i in A:
        if i == " " or i == ",":
            names.append(char_bucket)
        char_bucket.append(i)
            
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

    # for every nested list in the list...
    for i in GF:
        # compare the NRIC value
        search_NRIC()


        
            
        
    
    


def search_NRIC():
    if charnum > 9:
        return 0
    return
def search_Name():
    return

def gf_SetData():
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
    print(x.head)

