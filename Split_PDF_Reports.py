import PyPDF2
import os
import time
import shutil
import sys

######################################################################
# Split_PDF_Reports.py
#
# Splits PDF files based on bookmarks
#
# Parameters:
# 1. sourcePDFFile - Source PDF file to split
# 2. outputPDFDir - Output directory for split files
# 3. outputNamePrefix - Prefix to append to file names
# 4. deleteSourcePDF - Delete source PDF file after split (True/False)
######################################################################

# Helper class used to map pages numbers to bookmarks
class BookmarkToPageMap(PyPDF2.PdfFileReader):

    def getDestinationPageNumbers(self):
        def _setup_outline_page_ids(outline, _result=None):
            if _result is None:
                _result = {}
            for obj in outline:
                if isinstance(obj, PyPDF2.pdf.Destination):
                    _result[(id(obj), obj.title)] = obj.page.idnum
                elif isinstance(obj, list):
                    _setup_outline_page_ids(obj, _result)
            return _result

        def _setup_page_id_to_num(pages=None, _result=None, _num_pages=None):
            if _result is None:
                _result = {}
            if pages is None:
                _num_pages = []
                pages = self.trailer["/Root"].getObject()["/Pages"].getObject()
            t = pages["/Type"]
            if t == "/Pages":
                for page in pages["/Kids"]:
                    _result[page.idnum] = len(_num_pages)
                    _setup_page_id_to_num(page.getObject(), _result, _num_pages)
            elif t == "/Page":
                _num_pages.append(1)
            return _result

        outline_page_ids = _setup_outline_page_ids(self.getOutlines())
        page_id_to_page_numbers = _setup_page_id_to_num()

        result = {}
        for (_, title), page_idnum in outline_page_ids.items():
            result[title] = page_id_to_page_numbers.get(page_idnum, '???')
        return result

def main(arg1, arg2, arg3, arg4):
    ################
    # Main Program #
    ################
    #Set parameters
    sourcePDFFile = arg1
    outputPDFDir = arg2
    outputNamePrefix = arg3
    deleteSourcePDF = arg4
    targetPDFFile = 'temppdfsplitfile.pdf' # Temporary file

    if outputPDFDir:
        if os.path.exists(outputPDFDir):
            pass
        else:
            os.mkdir(outputPDFDir)

    print('Parameters:')
    print(sourcePDFFile)
    print(outputPDFDir)
    print(outputNamePrefix)
    print(targetPDFFile)

    #Verify PDF is ready for splitting
    while not os.path.exists(sourcePDFFile):
        print('Source PDF not found, sleeping...')
        #Sleep
        time.sleep(10)
    
    if os.path.exists(sourcePDFFile):
        print('Found source PDF file')
        #Copy file to local working directory
        shutil.copy(sourcePDFFile, targetPDFFile)

        #Process file
        pdfFileObj2 = open(targetPDFFile, 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj2)
        pdfFileObj = BookmarkToPageMap(pdfFileObj2)

        #Get total pages
        numberOfPages = pdfReader.numPages
        print('PDF # Pages: ' + str(numberOfPages))

        i = 0
        newPageNum = 0
        prevPageNum = 0
        newPageName = ''
        prevPageName = ''
        for p,t in sorted([(v,k) for k,v in pdfFileObj.getDestinationPageNumbers().items()]):
            template = '%-5s  %s'
            print (template % ('Page', 'Title'))
            print (template % (p+1,t))
        
            newPageNum = p + 1
            newPageName = t

            if prevPageNum == 0 and prevPageName == '':
                print('First Page...')
                prevPageNum = newPageNum
                prevPageName = newPageName
            else:
                if newPageName:
                    print('Next Page...')
                    pdfWriter = PyPDF2.PdfFileWriter()
                    page_idx = 0 
                    for i in range(prevPageNum, newPageNum):
                        pdfPage = pdfReader.getPage(i-1)
                        pdfWriter.insertPage(pdfPage, page_idx)
                        print('Added page to PDF file: ' + prevPageName + ' - Page #: ' + str(i))
                        page_idx+=1

                    pdfFileName = outputNamePrefix + str(str(prevPageName).replace(':','_')).replace('*','_') + '.pdf'
                    fullFilePath = os.path.join(outputPDFDir, pdfFileName)
                    pdfOutputFile = open(fullFilePath, 'wb')
                    pdfWriter.write(pdfOutputFile)
                    pdfOutputFile.close()
                    print('Created PDF file: ' + fullFilePath)

            i = prevPageNum
            prevPageNum = newPageNum
            prevPageName = newPageName

        #Split the last page
        print('Last Page...')
        pdfWriter = PyPDF2.PdfFileWriter()
        page_idx = 0 
        for i in range(prevPageNum, numberOfPages + 1):
            pdfPage = pdfReader.getPage(i-1)
            pdfWriter.insertPage(pdfPage, page_idx)
            print('Added page to PDF file: ' + prevPageName + ' - Page #: ' + str(i))
            page_idx+=1
        
        pdfFileName = outputNamePrefix + str(str(prevPageName).replace(':','_')).replace('*','_') + '.pdf'
        fullFilePath = os.path.join(outputPDFDir, pdfFileName)
        pdfOutputFile = open(fullFilePath, 'wb')
        pdfWriter.write(pdfOutputFile)
        pdfOutputFile.close()
        print('Created PDF file: ' + fullFilePath)

        pdfFileObj2.close()

        # Delete temp file
        os.unlink(targetPDFFile)

        if newPageName:
            if deleteSourcePDF == True or deleteSourcePDF == "True":
                os.unlink(sourcePDFFile)

if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4])
