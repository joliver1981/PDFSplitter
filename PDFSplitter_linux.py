import PyPDF2
import os
import time
import shutil
import sys
import argparse

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

def main(arg1, arg2, arg3, arg4, arg5):
    ################
    # Main Program #
    ################
    #Set parameters
    sourcePDFFile = arg1
    outputPDFDir = arg2
    outputNamePrefix = arg3
    deleteSourcePDF = arg4
    majorChaptersOnly = arg5
    targetPDFFile = 'temppdfsplitfile.pdf' # Temporary file

    #make sure output directory has slash at the end otherwise it does not path correctly
    if outputPDFDir[-1] is not "/":
        outputPDFDir = outputPDFDir + "/"

#I'm commenting it out, because we're running on Linux.
#    if outputPDFDir:
#        # Append backslash to output dir if necessary
#        if not outputPDFDir.endswith('\\'):
#            outputPDFDir = outputPDFDir + '\\'

    print('Parameters:')
    print(sourcePDFFile)
    print(outputPDFDir)
    print(outputNamePrefix)
    print(targetPDFFile)
    print(majorChaptersOnly)

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
        all_chapters = pdfFileObj.getDestinationPageNumbers().items()

        if majorChaptersOnly:
            chapter_list_used = []
            for c in all_chapters:
                chapter_title_strip = str(c[0]).strip()
                should_add_chapter = True
                for i in range(len(chapter_title_strip)):
                    # if chapter title is less than 3 characters
                    if len(chapter_title_strip) < 2:
                        break
                    # if reached end
                    if i > len(chapter_title_strip) - 2:
                        break
                    # look for pattern of digit, line or period, digit indicating a sub chapter
                    if chapter_title_strip[i].isdigit() and (chapter_title_strip[i+1] in [".", "-"]) and chapter_title_strip[i+2].isdigit():
                        should_add_chapter = False
                        break
                if should_add_chapter:
                    chapter_list_used.append(c)
                
        else:
            chapter_list_used = all_chapters

        for p,t in sorted([(v,k) for k,v in chapter_list_used]):
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
                        print('Added page to PDF file: ' + str(prevPageName) + ' - Page #: ' + str(i))
                        page_idx+=1

#                   pdfFileName = outputNamePrefix + str(str(prevPageName).replace(':','_')).replace('*','_') + '.pdf'
                    # I've added "/" as a character to replace, because, again, we're running on Linux
                    pdfFileName = outputNamePrefix + str(str(prevPageName).replace(':','_')).replace('*','_').replace('/',' ') + '.pdf'
                    pdfOutputFile = open(str(outputPDFDir) + str(pdfFileName), 'wb')
                    pdfWriter.write(pdfOutputFile)
                    pdfOutputFile.close()
                    print('Created PDF file: ' + outputPDFDir + pdfFileName)

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
        
#       pdfFileName = outputNamePrefix + str(str(prevPageName).replace(':','_')).replace('*','_') + '.pdf'
        # I've added "/" as a character to replace, because, again, we're running on Linux
        pdfFileName = outputNamePrefix + str(str(prevPageName).replace(':','_')).replace('*','_').replace('/',' ') + '.pdf'
        pdfOutputFile = open(outputPDFDir + pdfFileName, 'wb')
        pdfWriter.write(pdfOutputFile)
        pdfOutputFile.close()
        print('Created PDF file: ' + outputPDFDir + pdfFileName)

        pdfFileObj2.close()

        # Delete temp file
        os.unlink(targetPDFFile)

        if newPageName:
            if deleteSourcePDF == True or deleteSourcePDF == "True":
                os.unlink(sourcePDFFile)

if __name__ == "__main__":
    #Added argparse to deal with CL arguments and set defaults
    parser = argparse.ArgumentParser()
    parser.add_argument("input", help="input PDF file")
    parser.add_argument("output", nargs="?", help="output directory for the split   files", default="./Output/")
    parser.add_argument("prefix", nargs="?", help="split files' prefix", default="")
    parser.add_argument("delete", nargs="?", help="delete the original file? (True/False)", default=False)
    parser.add_argument("majorBookmarksOnly", nargs="?", help="Only output those chapters without - in the title", default=False)
    args = parser.parse_args()
    main(args.input, args.output, args.prefix, args.delete, args.majorBookmarksOnly)
