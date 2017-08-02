#!/usr/bin/env python3

"""
This program takes a known TSM reports PDF product
and parses it out to individual pdf files.
Use the TSM Reports option.
Don't use the TSM Export PDF option, it is not supported.
"""
# Jonathan McDonald 7/26/2017 11:19AM
# iteration 5

import PyPDF2
import copy
import re
import datetime
import os

# Number of new files created
numFiles = 0

# Date time stamp
now = datetime.datetime.now()

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

# Currently only prints first page
# TODO: Bug; some reports have static variables that span two or more pages, workaround is checking the error output
multiPage = False  # Forced as False until TODO is completed, then set True
# TODO: overriding print first page only
# paperSaving = boolean(input('Save paper by only printing first page of each record? (True or False): '))

# Static Variables #
# OpenningHeader
# FOR OFFICIAL USE ONLYTSM TRIAL CARD RECORD REPORT
openHeader = 'FOR OFFICIAL USE ONLYTSM TRIAL CARD RECORD REPORT'
# checkP1
# Generated on:12/31/1999
checkP1 = 'Generated on:'
# checkP2
# Trial Card:ABC0000DE-FG000101
checkP2 = 'Trial Card:'
# checkP4
# DiscoveredScreeningStatus DateStatusAction TakenAuthority
# reference values
checkP4 = 'DiscoveredScreeningStatus DateStatusAction TakenAuthority'
# checkP5
# Equipment:Some piece of gear
checkP5 = 'Equipment:'
# checkP6
# Rationale:
checkP6 = 'Rationale:'
# checkP7
# Transferred fromTransferred toCompartmentOriginatorPOCClosing Authority
# reference values
checkP7 = 'Transferred fromTransferred toCompartmentOriginatorPOCClosing Authority'
# checkP8
# ReferencesSWBSKey EventAssoc. Doc.Offsite 3
# reference values
checkP8 = 'ReferencesSWBSKey EventAssoc. Doc.Offsite 3'
# checkP9
# Amplifier
checkP9 = 'Amplifier'
# checkP10
# Narrative:The thing is not right.
checkP10 = 'Narrative:'
# checkP11
# Dispute/CommentsWorkflow HistoryWF NameStepTask \#Assigned ToCompleted byStartedDueCompleted
# reference values
checkP11 = 'Dispute/CommentsWorkflow HistoryWF NameStepTask'
# ClosingHeader
# FOR OFFICIAL USE ONLYpage1of2DIMSRecordAll
closingHeader = 'FOR OFFICIAL USE ONLYpage'

# Open the Error Log
try:
    fileFound = open(myPDFname + '_errorLog.txt', 'r')  # read to see if file exist
    fileFound.close()
    log_A_N = False
    while log_A_N is False:
        useLoose = input('Old error log found. Append to existing or Write New?\n(Enter "A" or "N"): ')  # Ask to use old or write new
        if useLoose == 'A' or 'a':
            errOut = open(myPDFname + '_errorLog.txt', 'a')  # append to existing log file
            errOut.write('~~~Log Appended on ' + str(now) + ' ~~~\n')
            errOut.write('Trial Card number' + '    ' + 'page number\n')
            errOut.close()
            log_A_N = True
        elif useLoose == 'N' or 'n':
            errOut = open(myPDFname + '_errorLog.txt', 'w')  # write a new log file
            errOut.write('~~~Log Created on ' + str(now) + ' ~~~\n')
            errOut.write('Trial Card number' + '    ' + 'page number\n')
            errOut.close()
            log_A_N = True
        else:
            print('Invalid log command. Enter either an [A] or a [N]')
except:
    errOut = open(myPDFname + '_errorLog.txt', 'w')  # write a new log file
    errOut.write('~~~Log Created on ' + str(now) + ' ~~~\n')
    errOut.write('Trial Card number' + '    ' + 'page number\n')
    errOut.close()
    print('New error log created')

# Open the Completed Log
try:
    fileFound = open(myPDFname + '_completeLog.txt', 'r')  # read to see if file exist
    fileFound.close()
    log_A_N = False
    while log_A_N is False:
        useLoose = input('Old completed log found. Append to existing or Write New?\n(Enter "A" or "N"): ')  # Ask to use old or write new
        if useLoose == 'A' or 'a':
            compOut = open(myPDFname + '_completeLog.txt', 'a')  # append to existing log file
            compOut.write('~~~Log Appended on ' + str(now) + ' ~~~\n')
            compOut.write('Trial Card number\n')
            compOut.close()
            log_A_N = True
        elif useLoose == 'N' or 'n':
            compOut = open(myPDFname + '_completeLog.txt', 'w')  # write a new log file
            compOut.write('~~~Log Created on ' + str(now) + ' ~~~\n')
            compOut.write('Trial Card number\n')
            compOut.close()
            log_A_N = True
        else:
            print('Invalid log command. Enter either an [A] or a [N]')
except:
    compOut = open(myPDFname + '_completeLog.txt', 'w')  # write a new log file
    compOut.write('~~~Log Created on ' + str(now) + ' ~~~\n')
    compOut.write('Trial Card number\n')
    compOut.close()
    print('New completion log created')

# with open('_errorLog.txt', 'a') as errorOut:
# with open('_completeLog.txt', 'a') as compOut:

with open(myPDFname + '_errorLog.txt', 'a') as errOut:
    with open(myPDFname + '_completeLog.txt', 'a') as compOut:

        # Begin running through the pages to extract pages for printing #
        for p in range(0, pdfPages):
            multiPage = False
            curPageObj = pdfReader.getPage(p)  # goto page number p
            pageObj = copy.copy(curPageObj)  # reassign to preserve
            # pageObj = pdfReader.getPage(p)  # goto page number p
            pageStrTxt = pageObj.extractText()  # extract a text string from page
            # Get the trial card number off of the current page with reg-ex
            # Trial Card:ABC0000DE-FG000101
            # literal string "Trial Card:"
            # two to four letters (\w){2,4}
            # four numbers (\d){4}
            # two to four letters (\w){2,4} and optional one number (\d)?
            # one hyphen -
            # two letters (\w){2}
            # six numbers (\d){6}
            reCKp2 = re.compile(r'Trial Card:(\w){2,4}(\d){4}(\w){2,4}(\d)?-(\w){2}(\d){6}')
            moCKp2 = reCKp2.search(pageStrTxt)  # returned page string
            if moCKp2 is None:
                # move to next page, this will be the file name and is required
                errOut.write('Trial Card number not found on page.\n')
                errOut.write(str(moCKp2) + '    ' + str(p) + '\n')
                continue  # Could be a break point to exit with an error
            else:
                fullCKp2 = moCKp2.group()  # matched object returned
                curCKp2 = fullCKp2[11:]  # slice out the literal string "Trial Card:"
                Page1_TCnum = curCKp2  # this is the trial card number, used as a reference
                outFileName = curCKp2 + '.pdf'  # this is the output file name

            # checkP3
            # Star*Priority1SafetyS
            # four letters (\w){4} with optional "*" (\W)?
            # eight letters (\w){8} with one number (\d)
            # six letters (\w){6} with optional "S" (\w)?
            reCKp3 = re.compile(r'(\w){4}(\W)?(\w){8}(\d)(\w){6}(\w)?')
            moCKp3 = reCKp3.search(pageStrTxt)  # returned page string
            if moCKp3 is None:
                print('Star Priority Safety check point not found on page.')
                # move to next page
                checkP3 = False  # should not be needed, here as a safety catch
                errOut.write('Star Priority Safety check point not found on page.\n')
                errOut.write(str(curCKp2) + '    ' + str(p) + '\n')
                continue
            else:
                curCKp3 = moCKp3.group()  # matched object returned
                checkP3 = True

            # Set to True to look for False
            fullPage = True
            # Look for a full page of TC Info
            if openHeader not in pageStrTxt:
                fullPage = False
            if checkP1 not in pageStrTxt:
                fullPage = False
            if checkP2 not in pageStrTxt:
                fullPage = False
            if not(checkP3 is True):
                fullPage = False
            if checkP4 not in pageStrTxt:
                fullPage = False
            if checkP5 not in pageStrTxt:
                fullPage = False
            if checkP7 not in pageStrTxt:
                fullPage = False
            if checkP8 not in pageStrTxt:
                fullPage = False
            if checkP9 not in pageStrTxt:
                fullPage = False
            if checkP10 not in pageStrTxt:
                fullPage = False
            if checkP11 not in pageStrTxt:
                fullPage = False
            if closingHeader not in pageStrTxt:
                fullPage = False

            if fullPage is False:
                # print('Partial Page found, skip to next page.')
                # move to next page
                # TODO: overriding print first page only
                # Page Range
                multiPage = True
                # multiPage = False
                # continue
                errOut.write(str(curCKp2) + '    ' + str(p) + '\n')

            # TODO: overriding print first page only, build a multi page export to pdf
            # If current card continues to next page would be throwaway loop just to increment page number
            # Page Range

            PageN_TCnum = Page1_TCnum
            n = copy.copy(p)
            print('Page start number: ' + str(p))
            print('Multi page start number: ' + str(n))

            # TODO: overriding print first page only
            while multiPage is True:
                n = n + 1
                if n > pdfPages:
                    # Exit page range
                    n = n - 1  # one page to far
                    print('out of range, next page exceeded total pages: ' + str(n))
                    multiPage = False
                else:
                    curNPageObj = pdfReader.getPage(n)  # goto next page number n
                    pageNObj = copy.copy(curNPageObj)  # reassign to preserve
                    # pageNObj = pdfReader.getPage(n)  # goto page number p
                    pageNStrTxt = pageNObj.extractText()  # extract a text string from page

                    # Get the trial card number off of the current page with reg-ex
                    # Trial Card:ABC0000DE-FG000101
                    # literal string "Trial Card:"
                    # two to four letters (\w){2,4}
                    # four numbers (\d){4}
                    # two to four letters (\w){2,4} and optional one number (\d)?
                    # one hyphen -
                    # two letters (\w){2}
                    # six numbers (\d){6}
                    reNCKp2 = re.compile(r'Trial Card:(\w){2,4}(\d){4}(\w){2,4}(\d)?-(\w){2}(\d){6}')
                    moNCKp2 = reNCKp2.search(pageNStrTxt)  # returned page string
                    if moNCKp2 is None:
                        print('Trial Card number not found on following page.\n')
                        # Exit page range
                        multiPage = False
                    else:
                        fullNCKp2 = moNCKp2.group()  # matched object returned
                        curNCKp2 = fullNCKp2[10:]  # slice out the literal string "Trial Card:"
                        PageN_TCnum = curNCKp2  # this is Trial Card number on the next page

                if not(PageN_TCnum == Page1_TCnum):
                    print('Prior page Trial Card number not found on this page.\n')
                    # Exit page range
                    n = n - 1  # one page to far
                    multiPage = False
                else:
                    # Un-captured Error
                    pass

                pageRangeUpper = n

            # TODO: overriding print first page only, pageRangeUpper is set in above block
            # that is out of current scope
            pageRangeUpper = n
            print('Page range upper number: ' + str(pageRangeUpper))

            pdfWriter = PyPDF2.PdfFileWriter()  # Make pdf write too object in memory

            if pageRangeUpper == p:
                # only copy one page
                print(outFileName)
                pageObj = pdfReader.getPage(p)
                pdfWriter.addPage(pageObj)
                pdfOutputFile = open(outFileName, 'wb')  # open destination file
                pdfWriter.write(pdfOutputFile)
                pdfOutputFile.close
                # Increment file counter
                numFiles = numFiles + 1
                # move to next page
                compOut.write(str(curCKp2) + '\n')
                continue

            elif pageRangeUpper < p:
                # out of range error, cart is before the horse
                break
            elif pageRangeUpper > p:
                if pageRangeUpper > pdfPages:
                    # out of range error, last page exceeded parent doc final page
                    break
                elif pageRangeUpper <= pdfPages:
                    if not (os.path.isfile(outFileName)):
                        # TODO: overriding print first page only
                        if paperSaving is True:
                            # only copy one page
                            # TODO: Check if file exist
                            pageObj = pdfReader.getPage(p)
                            pdfWriter.addPage(pageObj)
                            pdfOutputFile = open(outFileName, 'wb')  # open destination file
                            pdfWriter.write(pdfOutputFile)
                            pdfOutputFile.close
                            # Increment file counter
                            numFiles = numFiles + 1
                            # move to next page
                            errOut.write(str(curCKp2) + '    ' + str(p) + ' completed\n')
                            compOut.write(str(curCKp2) + '\n')
                            continue
                        else:
                            # copy the page range
                            start = p
                            end = pageRangeUpper
                            for pageInc in range(start, end):
                                pageObj = pdfReader.getPage(pageInc)
                                pdfWriter.addPage(pageObj)

                            pdfOutputFile = open(outFileName, 'wb')  # open destination file
                            pdfWriter.write(pdfOutputFile)
                            pdfOutputFile.close
                            # Increment file counter
                            numFiles = numFiles + 1
                            # move to next page
                            errOut.write(str(curCKp2) + '    ' + str(p) + ' completed\n')
                            compOut.write(str(curCKp2) + '\n')
                            continue
                    else:
                        # File exist, don't overwrite the first file
                        continue
                else:
                    # Un-captured Error
                    continue
            else:
                # Un-captured Error
                continue


# Completed
print('Finished copying pages to new pdf files')
print('Number of new files created: ' + str(numFiles))

# End Of Line
hold = input('Press any key to exit.')

# End Of Line
print('goodbye')
