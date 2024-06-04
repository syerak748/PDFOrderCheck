import fitz  # PyMuPDF
import re
import pandas as pd
from math import isnan
import datetime
pdffilePath = "/Users/rxnkshitij748/Downloads/SAMPLEPDF.pdf" #PDF
samplerefsheetpath = "/Users/rxnkshitij748/Downloads/SampleReferenceSheet copy.xlsx" #REFERENCE SHEET
ErrorLogFile = "logErrors.txt" #LOG FILE


#------------------------------------------------EXCEL CUSTOM PARSING----------------------------------------------------


def parseExcelToListDict(refSheet = samplerefsheetpath):#parses relevant columns and extracts data as list of dicts. CONFIGURABLE
    df = pd.read_excel(samplerefsheetpath, header=2, usecols=['Form Name', 'Form Number', 'INDEX']) #configure here
    parsedListDict = df.to_dict(orient="records")
    return parsedListDict

#formIDlist = ["Cover Letter", '1234A', '1236B', '117109', '411379', '112065', '115465', 'ICC21-117101', '117101', '117101-CA', '117101-FL']
def constructFormIDList(refSheet = samplerefsheetpath):#constructs the id list, only one condition, i.e, if "Form Number == 'n/a' then see "Form Name"
    parsedReference = parseExcelToListDict(refSheet)
    FormIDlist = []
    for row in parsedReference:
        if row['Form Number'] == "n/a" or (type(row['Form Number']) == float):
            if isnan(row['Form Number']) or row['Form Number'] == "n/a":
                FormIDlist.append(str(row['Form Name']))
        else:
            FormIDlist.append(str(row['Form Number']))
    return FormIDlist

def constructFormIDListIndexed(refSheet = samplerefsheetpath):#constructs the id list, only one condition, i.e, if "Form Number == 'n/a' then see "Form Name"
    parsedReference = parseExcelToListDict(refSheet)
    IndexedFormIDlist = []
    for row in parsedReference:
        if row['Form Number'] == "n/a" or (type(row['Form Number']) == float):
            if isnan(row['Form Number']) or row['Form Number'] == "n/a":
                IndexedFormIDlist.append({'FormID' : str(row['Form Name']), 'Index' : row['INDEX']})
        else:
            IndexedFormIDlist.append({'FormID' : str(row['Form Number']), 'Index' : row['INDEX']})
    return IndexedFormIDlist

def constructFormIDDictIndexed(refSheet = samplerefsheetpath):#constructs the id list, only one condition, i.e, if "Form Number == 'n/a' then see "Form Name"
    parsedReference = parseExcelToListDict(refSheet)
    IndexedFormIDdict = {}
    for row in parsedReference:
        if row['Form Number'] == "n/a" or (type(row['Form Number']) == float):
            if isnan(row['Form Number']) or row['Form Number'] == "n/a":
                IndexedFormIDdict[str(row['Form Name'])] = row['INDEX']
        else:
            IndexedFormIDdict[str(row['Form Number'])] = row['INDEX']
    return IndexedFormIDdict

#------------------------------------------------PDF CUSTOM PARSING----------------------------------------------------

def extractReqdOutlines(pdffilePath=pdffilePath):  # extract suboutlines
    pdfDoc = fitz.open(pdffilePath)
    outlines = pdfDoc.get_toc()
    suboutlinesList = []
    for item in outlines:
        level, title, pageNo = item
        if level == 2:  # can change according to document.... lowest level of outlines
            suboutlinesList.append({"title": title, "level": level, "pageNo": pageNo})   
    return suboutlinesList

def printsubOutlines(suboutlines):  # testing function
    for suboutlineNo, suboutline in enumerate(suboutlines):
        print(f"Sno : {suboutlineNo} , title : {suboutline['title']}, PageNo : {suboutline['pageNo']}\n")
    return suboutlines

def filtersubOutlines(suboutlines, formIDlist = constructFormIDList(), Verification = True):  # filters so only one instance of a form_number
    filteredIDlist = [] # new list
    filterefmemorylist = [] #contains both filteredoutline and the formid corresponding
    visitedIDs = set() # to ensure same ID isnt caught and registered in two different outlines
    formRegexList = [re.compile(re.escape(form), re.IGNORECASE) for form in formIDlist] # regex to convert all reference sheet IDs to searchable form, case insensitive, and treats all special characters as literals
    '''
    print("Suboutlines titles:")
    for suboutline in suboutlines:
        print(suboutline['title'])

    print("\nForm IDs and their regex patterns:")
    for form, formRegex in zip(formIDlist, formRegexList):
        print(f"Form: {form}, Regex: {formRegex.pattern}")''' #DEBUGGER

    for suboutline in suboutlines:
        for form, formRegex in zip(formIDlist, formRegexList):
            if formRegex.search(suboutline['title']):#search for the regex pattern of the ID in the outline
                if form not in visitedIDs: # new ID
                    visitedIDs.add(form)
                    filteredIDlist.append(suboutline)
                    filterefmemorylist.append({"filteredoutline" : suboutline['title'], "pdfOrderedformid" : form})
                    break
                else:
                    #print(f"{suboutline['title']} removed\n")
                    pass
    if Verification:
        print("\nVerification:")
        for suboutlineno, suboutline in enumerate(filteredIDlist):
            print(f"Filtered Outline: {suboutline['title']} | {suboutlineno}")

    return filterefmemorylist

#------------------------------------------------------ORDER CHECKING----------------------------------------------------

def checkErrors(filteredWithMemory, IndexedFormIDList, FormIdIndexDict): #FUNCTION TO CHECK ERRORS IN ORDER OF FORMS ACCORDING TO THE EXCEL REFERENCE SHEET
    i = 0
    #open(ErrorLogFile, "w").close()
    print()
    errors = 0
    for i in range(0,len(filteredWithMemory)):
        if filteredWithMemory[i]['pdfOrderedformid'] == IndexedFormIDList[i]['FormID']:
            print(f"{filteredWithMemory[i]['pdfOrderedformid']} == {IndexedFormIDList[i]['FormID']} | {i} ")
            continue
        else:
            errors += 1
            print(f"{filteredWithMemory[i]['pdfOrderedformid']} != {IndexedFormIDList[i]['FormID']} | {i} ")
            with open(ErrorLogFile, "a") as f:
                logErrorMessage = f'Error in the order : {filteredWithMemory[i]['pdfOrderedformid']} should be at position {FormIdIndexDict[filteredWithMemory[i]['pdfOrderedformid']]}\n'
                f.write(logErrorMessage)
    return errors

def checkPageOrder(suboutlines): #FUNCTION TO CHECK ORDER OF PAGES ACCORDING TO KEY "PG" CASE INSENSITIVE
    pageOrderErrors = 0
    formPages = {}
    for suboutline in suboutlines:
        title = suboutline['title']
        match = re.match(r"(.*?)( pg\d+)", title, re.IGNORECASE) #CAN CHANGE KEY FOR PAGE NUMBER AS PER DOCUMENT
        if match:
            formId = match.group(1).strip() #matches group 1 that is formID from title
            pageNo = match.group(2).strip() #matches group2 from remaining title ,i.e, pg
            if formId not in formPages: #if the pg doesnt match, match returns None so only gets the multi page forms
                formPages[formId] = []
            formPages[formId].append(pageNo)
    
    
    '''
    print("Collected form pages:")
    for form_id, pages in form_pages.items():
        print(f"Form ID: {form_id}, Pages: {pages}")''' #DEBUGGER
    for formId, pages in formPages.items():
        sortedPages = sorted(pages, key=lambda x: int(re.search(r'\d+', x).group())) #sorts the pageNo list in dictionary of lists FormPages 
        if pages != sortedPages: #checks order actual wrt to sorted PageNos
            pageOrderErrors += 1
            with open(ErrorLogFile, "a") as f:
                f.write(f'\nPage order error: Pages for {formId} are not in order: {pages}\n')
    
    return pageOrderErrors

#---------------------------------------------------MAIN FUNCTION----------------------------------------------------

errorYes = True
def mainfn(ResetFile = False): # MAIN FUNCTION COMPILES AND RETURNS EVERYTHING; 
    if ResetFile:
        open(ErrorLogFile, "w").close()
    with open(ErrorLogFile,"a") as f:
        f.write(f'\nERROR LOGS FOR PDF -> {pdffilePath} at [{datetime.datetime.now()}] : \n')
    global errorYes
    formIDs = constructFormIDList()
    suboutlines = extractReqdOutlines()  # No changes here, using your original function
    filteredWithMemory = filtersubOutlines(suboutlines, formIDs, Verification=True)
    indexedFormIDList = constructFormIDListIndexed()
    indexedFormIDdict = constructFormIDDictIndexed()
    errors = checkErrors(filteredWithMemory, indexedFormIDList, indexedFormIDdict)
    pageOrderErrors = checkPageOrder(suboutlines)
    if errors == 0 and pageOrderErrors == 0:
        errorYes = False
        with open(ErrorLogFile,"a") as f:
            f.write("No Errors, pdf is in order\n")
        print("No Errors, pdf is in order")
    else:
        with open(ErrorLogFile,"a") as f:
            f.write(f'Total Errors: {errors + pageOrderErrors}\n')
        print(f'Total Errors: {errors + pageOrderErrors}')

if __name__ == "__main__": #TERMINAL RUN CONFIG
    mainfn(True) #CHANGE THIS TO FALSE TO SAVE PREVIOUS ERROR LOGS
    if errorYes:
        print(f'Errors recorded at {ErrorLogFile}')
    else:
        pass












'''TESTER
suboutlines = extractReqdOutlines()

printsubOutlines()
print()
newSubs, newsubswithmemory = filtersubOutlines(suboutlines)
indexedlist = constructFormIDListIndexed()

refParsed = parseExcelToListDict()
from pprint import pprint
pprint(refParsed)
pprint(newSubs)
testlist = constructFormIDList()
# Verify if 'FL Cover_117109_pg1' is included in the filtered list

'''

#table processing task
'''


'''