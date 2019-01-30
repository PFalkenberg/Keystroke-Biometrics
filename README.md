# What this program does

The xlsx file contains raw keystroke biometric data from 51 users typing a password 400 times each. This program uses a 
specified number of trials to create a template vector for each user and save these templates in their own spreadsheet. 
The remaining trials from each user are used as probe vectors and are tested against the templates to generate scores. In 
the case where the first 200 trials are used to generate template vectors, the remaining 200 for each user would be used as 
probes and generate a total of 520,200 scores. The program will calculate the false positive and false negative rates at all 
possible thresholds. The rates for specific threshold can be searched for using the GUI.

# Problems encountered

For this project, I used java with the apache poi library to read/write the xlsx file. This method worked effectively, until the 
program was near completion. As it turns out, apache poi is very memory hungry and the program would stall out and crash when trying 
to calculate FPRs and FRRs. By the time I encountered this problem, it would have been too costly to start over using a new method, so 
a workaround was put into place. As it is now, I had each step save the calculated values as csv files and then manually converted them 
to xlsx files for use in the next calculation. If I were to redo this project, I would definitely use csv files the whole way through.
