# Notes

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
      4. Ok this function now will return the rows that matches were found. If none found, just send back everything

4. match name
   1. try to match the whole name first
   2. think abt it: even if you try to search each segment, how confident will you be that that is the correct person? If you only search "Tan" and one person comes out, then ok thats fine. But also you can search "Fvsdn" and none come out..? In that case then go to the next name segment and try to find. If have multiple, then add the next name segment and try to narrow down to 1 record.
      1. **may not be the case. What if this guy put "george", but his bank only has chinese name. But at the same time there rly is another guy called "george". Then the system will think, oh ok i found him.** 

   3. Ok so we're only matching exact full name. Let's test.
      1. Hey! Just because you single out some in NRIC, you still missed out some that only match name. Don't just send the matched nrics only! Refactor the code. `search_Name` and `search_NRIC` 
         1. nope revert. i create a "duplicate" list to handle multiples.
         2. Just check name regardless, you don't save a lot of time anyway.

5. If in the end you still end up with multiple, then check back on the googleform if there have been multiple of this name.
   1. if so add them to a NEW list. Called "paid for other ppl"
   2. if not, just flag this guy out as paid too many times.

6. printing it back out on excel
   1. it's ok to overwrite. This is because we want the list to keep updating. Actually, don't overwrite, make a new file each time, with naming convention including the date.
      1. This is good for archival purposes
7. The logic shouldn't even need to reach checkAgainstGF actually. once checkAgainstBank is 0, it appends alr.
