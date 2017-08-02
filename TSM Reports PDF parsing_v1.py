#!/usr/bin/env python3

"""
This program takes a known TSM reports PDF product
and parses it out to individual pdf files
"""
# Jonathan McDonald 7/25/2017 4:06PM
# iteration 2.5

import PyPDF2
import copy
import re

# Number of new files created
numFiles = 0

# TSM_ExportOption Don't use this report .pdf
# TSM_ReportsOption this is the report to use .pdf
myPDFname = input('Enter PDF file name:')  # file name
# TODO: overriding print first page only
# paperSaving = boolean(input('Save paper by only printing first page of each record? (True or False): '))
pdfFileObj = open(myPDFname, 'rb')  # file handle

pdfReader = PyPDF2.PdfFileReader(pdfFileObj)  # read the pdf into workspace
pdfPages = pdfReader.numPages
print('Page Count : ' + str(pdfPages))

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

# Begin running through the pages to extract pages for printing #
for p in range(0, pdfPages):
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
        print('Trial Card number not found on page.')
        # move to next page, this will be the file name and is required
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

    if not (fullPage is True):
        # print('Partial Page found, skip to next page.')
        # move to next page
        # TODO: overriding print first page only
        # Page Range
        multiPage = False
        continue

    # TODO: overriding print first page only, build a multi page export to pdf
    # If current card continues to next page would be throwaway loop just to increment page number
    # Page Range
    multiPage = False  # Forced as False until TODO is completed, the set True
    PageN_TCnum = Page1_TCnum
    n = copy.copy(p)

    # TODO: overriding print first page only
    while multiPage is True:
        n = n + 1
        if n > pdfPages:
            # Exit page range
            n = n - 1  # one page to far
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
                print('Trial Card number not found on following page.')
                # Exit page range
                multiPage = False
            else:
                fullNCKp2 = moNCKp2.group()  # matched object returned
                curNCKp2 = fullNCKp2[10:]  # slice out the literal string "Trial Card:"
                PageN_TCnum = curNCKp2  # this is Trial Card number on the next page

        if not(PageN_TCnum == Page1_TCnum):
            print('Prior page Trial Card number not found on this page.')
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
        continue

    elif pageRangeUpper < p:
        # out of range error, cart is before the horse
        break
    elif pageRangeUpper > p:
        if pageRangeUpper > pdfPages:
            # out of range error, last page exceeded parent doc final page
            break
        elif pageRangeUpper <= pdfPages:
            # TODO: overriding print first page only
            if paperSaving is True:
                # only copy one page
                pageObj = pdfReader.getPage(p)
                pdfWriter.addPage(pageObj)
                pdfOutputFile = open(outFileName, 'wb')  # open destination file
                pdfWriter.write(pdfOutputFile)
                pdfOutputFile.close
                # Increment file counter
                numFiles = numFiles + 1
                # move to next page
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
