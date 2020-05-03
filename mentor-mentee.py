import pandas as pd

# open the file
tees = pd.ExcelFile(mentees.xlsx)
tors = pd.ExcelFile(mentors.xlsx)

# get the first sheet as an object
mentees = tees.parse(0)
mentors = tors.parse(0)

# get the first column as a list you can loop through
# where the is 0 in the code below change to the row or column number you want    
eerow = mentees.iloc[0,:]
eecolumn = mentees.iloc[:,0]
print(eerow,eecolumn)

# get the first row as a list you can loop through
orrow = mentors.iloc[0,:]
orcolumn = mentors.iloc[:,0]
print(orrow,orcolumn)
