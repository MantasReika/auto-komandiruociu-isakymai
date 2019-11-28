#import os
#import json
#import copy
#import openpyxl
#from docx import Document
from pprint import pprint

from wordbuilder_procedures import getCountry, getGalininkaVarda, getXlsxData, getDienpinigiai, indexParagrahps, indexRuns, createSecondmentDocument, loadJsonData
from wordbuilder_procedures import Secondment
from wordbuilder_procedures import Constants


    
def main():
    c = Constants()
    doc = Document(c.TEMPLATE_NAME)
    try:
        os.mkdir(c.DOCUMENTS_DIR)
    except FileExistsError:
        pass

    try:
        os.mkdir(c.FAILED_DOCUMENTS_DIR)
    except FileExistsError:
        pass

    records, unfit_records = getXlsxData("isakymai.xlsx", c.GALININKO_LINKSNIAI, c.INVALID_GA, c.INVALID_SPEC, c.COUNTRIES)
    
    paragraphIdx = indexParagrahps(c)
    runIdx = indexRuns(c, paragraphIdx)
    
    for each_record in records:
        createSecondmentDocument(each_record, copy.deepcopy(doc), paragraphIdx, runIdx, c, True)
   
    for each_record in unfit_records:
        createSecondmentDocument(each_record, copy.deepcopy(doc), paragraphIdx, runIdx, c, False)

    a = input("Finished...")
    


##
##    
##    pprint(paragraphIdx, width=1)
##    print("\n")
##    pprint(readIdx, width=1)
            
if __name__ == "__main__":
    main()
