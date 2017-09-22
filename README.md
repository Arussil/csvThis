# csvThis

### TL;DR; Description
A little project that, for now, convert xlsx files to CSV.

### Requirements
Python 3.4+

Openpyxl 2.4.8+

Requirements are more loose than it seems, they are just the version i had up when i created this.

### Description
Don't you hate when client just hand you a shitty format for some data to import? Even if you ask specifically for a .csv you still get spreadsheet in multiple exel files/sheets.

That's what csvThis! (! for emphasis) is for, it just get a .xlsx (for now, i'll try to add more format) and convert it to a .csv with the delimiter you wants.
I created it to help a collegue with a really bad formatted .xlsx in a couple of minute,then found out we have MORE xlsx to handle and so i worked on it some more to accept some external parameters (and do some test with the argparse and pathlib libraries)

In your face Excel!

### Things to do
A lot, more source files, better cleaning, custom function for add your own validation, create more than one .csv (for creating smaller import files in case you are stuck with bad hosting), a way to read all excel file sheet and build a .csv file for each one and (soon as i need it for a project) a way to take more than one file with same headers and get only one .csv.
