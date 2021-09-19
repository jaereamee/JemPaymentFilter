1. first she put the relevant data into the correct sheets
2. get data from gf and bank into a nested list (pandas df). Row-major.
   1. try again with pandas
      1. **Ok**
   2. I need to ensure all are string so can be reflected properly when i print it back
3. compare NRIC
   1. I need to know which is NRIC
      1. **just take fixed column? In the UI let the user say which is the NRIC column.**
   2. i can't take length of those NaN values, cos they are float type.
      1. so if Nan, put blank?
      2. I alr set dtype to str but it's still float for Nan
      3. use `df.replace('nan','')`?
      4. No. Use `pandas.df.fillna('')`
         1. **Yes it works!**
   3. now compare it against the bank records
      1. why i can't i pull out 1 individual bank record? Getting KeyError.
         1. Bank.head() works so the df does make it into the function
         2. This is becuase it assumes the 1st row as the header, this is bad cos it changes with each time i paste smt
         3. can consider `df.iloc[]` it still ignores the header. What if this is impt?
            1. **nvm assume it is not first**
      2. ok i can use `base.find(key)`. 
         1. I am getting false positives cos those nan values become blank!
            1. Try getting rid of the nan to blank replacer
               1. no doesnt work
            2. Its cos the searching bank records should be within the iterating thru googleform loop, if not it will keep
         2. **Ok i have found how to do match by nric, `.find` works** 
      3. Now i collect all these gf records in another df. Need to send it back to the main. If they follow the NRIC format, there should not be any duplicates. It should only send back 1 record.
         1. IF there's more than 1, make noise. Add to the "flag" column "multiple NRIC records 
      4. Ok this function now will return the rows that matches were found.