#!/usr/bin/env python3

"""
This program takes a known TSM reports PDF product
and parses it out to individual pdf files.
Use the TSM Reports option.
Don't use the TSM Export PDF option, it is not supported.
Args:
Returns:
Raises:
"""
import PyPDF2
import os

# Get and test the PDF file
fileNotFound = True
while fileNotFound is True:
    myPDFname = input('Enter PDF file name:')  # file name
    try:
        pdfFileObj = open(myPDFname, 'rb')  # file handle
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)  # read the pdf into workspace
        pdfPages = pdfReader.numPages
        # print('Page Count : ' + str(pdfPages))
        fileNotFound = False
    except:
        print('File not found, please re-enter the file name.\nFile must be in the Python root directory.')

page = input('What page do you want to read? ')
p = int(page) - 1 # pdf document starts at 0

pdfBlob = open('pdfBlob_page_' + page +'.txt', 'w')  # write a new log file
pdfBlob.close()

with open('pdfBlob_page_' + page +'.txt', 'a') as pdfBlob:
    curPageObj = pdfReader.getPage(p)  # goto page number p
    pageStrTxt = curPageObj.extractText()  # extract a text string from page
    pdfBlob.write(myPDFname + 'pdf text blob page: ' + page + '\n')
    pdfBlob.write(pageStrTxt)

# Completed
print('Finished copying extracted page text to the exported text file.')

# End Of Line
hold = input('Press any key to exit.')

# End Of Line
print('Good Bye')
