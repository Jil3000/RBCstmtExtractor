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

from PyPDF2 import PdfFileReader

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
csvFilename = ""

######################
## Open the File    ##
######################

rawFile = open(pdfFilepath, 'rb')
PDFstatement = PdfFileReader(rawFile)
if PDFstatement.isEncrypted:
    PDFstatement.decrypt("")
PDFMetaInformation = PDFstatement.getDocumentInfo() # TODO: delete if not needed
numOfPages = PDFstatement.getNumPages() # TODO: delete if not needed

# with open(pdfFilepath, 'rb') as f:
#     pdf = PdfFileReader(f)
#     if pdf.isEncrypted:
#         pdf.decrypt("")
#     print(pdf.getNumPages())
    # else:
    #     print("not encrypted")

    # information = pdf.getDocumentInfo()
    # number_of_pages = pdf.getNumPages()





######################
## Driver           ##
######################

# create filename for csv
csvFilename = BANK + "-" + CC_TYPE + "-" + DOC_TYPE + "-" + DOC_MONTH + "-" + DOC_YEAR + ".csv"

# extract all text:
pageObj = PDFstatement.getPage(0)

rawText = pageObj.extractText()
print(rawText)



# close file
rawFile.close()