import os
from openpyxl import Workbook #lib for managing excel https://openpyxl.readthedocs.io/en/stable/index.html 
from openpyxl.utils import get_column_letter

cwd = os.getcwd()

def bank_GetData():
    
    ## open the workbook, create if doesn't exist
    BANK_PATH = cwd+"\bankStatements.xlsx"
    if not os.path.isfile(BANK_PATH): 
        wb=Workbook()
        ws1 = wb.active
        ws1.title = "Raw"
        create_wb = BANK_PATH
    
    
    
    #The bank statement always has `Total Debit Count` at the end of the document, so you can scan until there
    return

# i also need to read the info from google form
def gf_GetData():
    """it will be multiple column, with the key based on the fields of the form. Make this dynamic? Since google forms produces the first row as a fixed title name anyway"""
    return

# find the different formats of payment. PayNow has a format, Paylah, BankTransfer. If majority of the cases are paynow/paylah, then why not i automate that first. To do a bank transfer the patient must have emailed the clinic.
def bank_Searcher():
    """detects the format of a paynow/paylah and puts the relevant fields in a new column"""

    return

def bank_SplitName(A):
    """split the name into strings, separated by spaces and commas"""
    count=0
    name_components_bucket=[]
    names=["john","tan"] #these are the names of the patients
    for i in A:
        if i == " " or i == ",":
            names.append(char_bucket)
        char_bucket.append(i)
            
def masterFilter():
    """The core filtering logic"""



def search_NRIC():
    return
def search_Name():
    return

if __init__ == "__main__":
    ## lets deal with lists because we dunnid to do arithmetic, but we wna edit them alot.
    # first i need to get all the rows in bankStatements.xlsx and put it in a big list
    bank_GetData()

    # #also read the fields in the googleform. Search the first row for headers.

    #identify a paynow/paylah transaction in the long array, and pull the NRIC (if available) and each segment of their name.
        
    """compare these fields with every entry in the gf. Store in a list of possible matches, if more than one remains, unlock the next layer of search. 
    Do in this order..
    1. check if comments say "sinopharm", if so pull the NRIC
    2. check the first name segment, if multiple, continue
    3. check the rest of name segment until the "name components list" is empty.
    After singling out the match, port it, along with all the other info, into a new sheet called "Paid"
    """

    ## if at the end of this you still have multiple names, this means multiple payment. Put them in another sheet called "Multiple Payment"

    ## the rest of them put it in the "Not Paid" sheet