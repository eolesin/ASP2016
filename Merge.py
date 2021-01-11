# Import dependencies for use
import pandas as pd
import numpy as np

# Import a dummy list file of the ARISA lengths
dummy = pd.read_csv('dummy.csv')

# Prompt the user for file names to use in the merge
names = []

# Set new_name to something other than 'done'.
new_name = ''

# Start a loop that will run until the user enters 'done'.
while new_name != 'done':
    # Ask the user for a name.
    new_name = input("Please enter next CSV filename, or enter 'done': ")
    # Add the new name to our list.
    if new_name != 'done':
        names.append(new_name)

out_name = input("Please enter your OUTPUT CSV filename, or enter 'done': ")

# Show that the name has been added to the list.
print("Merging: ", names)

cnt = 0
dflist = []
for i in names:
    if names[0] ==i:
        dataframename = str(i) + str(cnt)
        list.append(dflist,dataframename)
        if dflist[cnt] == dataframename:
            res = pd.read_csv(i)
            resprim = pd.merge(dummy,res,on='Length', how='outer')
        cnt = cnt + 1
    if names[0] != i:
        dataframename = str(i) + str(cnt)
        list.append(dflist,dataframename)
        if dflist[cnt] == dataframename:
            res2 = pd.read_csv(i)
            resprim = pd.merge(resprim,res2,on='Length', how='outer')
        cnt = cnt + 1

r = resprim.loc[resprim[resprim.columns.difference(['Length'])].sum(axis=1) != 0]

r.to_csv(out_name, index=False)
