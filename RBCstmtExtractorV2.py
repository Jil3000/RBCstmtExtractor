######################################
##                                  ##
## RBC Statement Extractor          ##
## by Gillian Waters                ##
##                                  ##
######################################
# 
# A program to take an RBC PDF statement, pull out the 
# relevant data (each credit card charge or payment with
# details), and save it as a csv file.

######################
## Imports          ##
######################

# from PyPDF2 import PdfFileReader

# pdfminer.six
# https://pdfminersix.readthedocs.io/en/latest/tutorial/composable.html
from io import StringIO

from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfdocument import PDFDocument    # unused?
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser    # unused?

# regular expressions
import re

# csv files
import csv

######################
## Settings         ##
######################

# for naming the csv file:
BANK = "RBC"
CC_TYPE = "Mastercard"
DOC_TYPE = "Stmt"
DOC_MONTH = "Oct"
DOC_YEAR = "2021"

            # TODO: pull this info from the filename
            # supporting info: https://stackoverflow.com/questions/8384737/extract-file-name-from-path-no-matter-what-the-os-path-format

# PDF to extract:
pdfFilepath = r"C:\Users\Jil30\Downloads\Credit Card Statement-1621-2021-Oct-26.pdf"

#TODO: take a parameter for the file name and path

######################
## Classes          ##
######################

class Transaction:

    def __init__(self, txDate, postDate, description, txNum, amount):
        self.txDate = txDate
        self.postDate = postDate
        self.description = description
        self.txNum = txNum
        self.amount = amount

######################
## Extract All Text ##
######################

# this method is based on the pdfminer.six documentation referenced above, written myself
def extractRawText(filepath):
    output = StringIO()
    origFile = open(filepath, "rb")

    # the following lines are pretty much verbatim from the documentation.  Sets up the pdfminer
    # and processes the document.
    parser = PDFParser(origFile)
    doc = PDFDocument(parser)
    rsrcmgr = PDFResourceManager()
    device = TextConverter(rsrcmgr, output, laparams=LAParams())
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    for page in PDFPage.create_pages(doc):
        interpreter.process_page(page)

    # save text before closing IO string
    outputAsText = output.getvalue()

    # close up everything
    origFile.close()
    output.close()

    # return the text
    return outputAsText

# There is an alternate answer here, that I tested and it worked:
# https://stackoverflow.com/questions/26494211/extracting-text-from-a-pdf-file-using-pdfminer-in-python
# not using it because I don't understand it, it's more complicated than the above, and it's someone
# else's code.

######################
## Print in Console ##
######################

def printAllTxs(txObjList):

    # column formatting
    headerFormat = "{:^8} {:^8} {:^50} {:^25} {:^8}"
    colFormat = "{:<8} {:<8} {:<50} {:<25} {:>8}"

    # column names
    colNames = ["Tx Date", "Post Date", "Description", "Tx Number", "Amount"]

    # column data
    print(headerFormat.format("Tx Date", "Post Date", "Description", "Tx Number", "Amount"))
    for tx in txObjList:
        print(colFormat.format(tx.txDate, tx.postDate, tx.description, tx.txNum, tx.amount))

######################
## Eeeeeeeeeeeee    ##
######################

def countInvisibleLines(raw):
    newline = tab = ret = u2029 = 0

    for c in raw:
        if c == "\n":
            newline += 1
        elif c == "\t":
            tab += 1
        elif c == "\r":
            ret += 1
        elif c == "\u2029":
            u2029 += 1

    print("\\n: " + str(newline))
    print("\\t: " + str(tab))
    print("\\r: " + str(ret))
    print("\\u2029: " + str(u2029))



######################
## Driver           ##
######################

# create filename for csv
csvFilename = BANK + "-" + CC_TYPE + "-" + DOC_TYPE + "-" + DOC_MONTH + "-" + DOC_YEAR + ".csv"

# extract all text
rawText = extractRawText(pdfFilepath)

# separate out each transaction
txList = re.findall("[a-zA-Z]{3}[0-9]{2}[a-zA-Z]{3}[0-9]{2}.+?[0-9]{23}-?\$[0-9]+\.[0-9]{2}",rawText)

# break each transaction into the fields
txObjList = []
txListList = []
for tx in txList:
    txDate = tx[0:5]
    postDate = tx[5:10]
    txNum = "\'" + (re.findall("[0-9]{23}", tx[10:]))[0]
    amount = (re.findall("-?\$[0-9]+\.[0-9]{2}", tx[6:]))[0]
    descriptionEndLoc = tx.find(txNum)
    description = (re.findall(".+", tx[10:descriptionEndLoc]))[0]
    nextTxObj = Transaction(txDate, postDate, description, txNum, amount)
    txObjList.append(nextTxObj)
    nextTxList = [txDate, postDate, description, txNum, amount]
    txListList.append(nextTxList)
#TODO: I don't actually need the tx objects after all, I should clean those up if I can rewrite the print method without it.
#TODO: figure out pulling and adding year

# print(txListList)

# create empty csv file and file creator
# create the new file
csvFile = open(csvFilename, "x")
csvFile.close()
#TODO: I want to write to the file, but still throw an error if the file already exists.  Opening it twice is a workaround.
csvFile = open(csvFilename, 'w', newline = "")
# add data to csv
csvWriter = csv.writer(csvFile)

# concatenate it into a .csv
# CSVContents = ""
# add headers to the .csv file
colNames = ["Tx Date", "Post Date", "Description", "Tx Number", "Amount"]   #TODO: this is stolen from the method at the top.  Fix
csvWriter.writerow(colNames)
for tx in txListList:
    # add it to the .csv file
    csvWriter.writerow(tx)      #TODO: change to writerows() instead?
    print(tx)

# save the .csv as a file in the same location as the original file, using the new name



csvFile.close()

# Testing Only, Delete Later
# countInvisibleLines(pdfFilepath)
# print(rawText)
# print(txList)
# print(len(txList)) # 89           # s/b 89, statement has 89 on it
# print(len(txObjList)) # 89!
# printAllTxs(txObjList)