import os
import time
import datetime

from OCR_CommonUtil import getInputFileList, finalizedResultSet, writeJsonOutput, preProcessChecker, insertRecord, \
    validateExtractedRecord
from OCR_Config import ConfigParam
from OCR_Core_CustModel_ChatHorn import runChatHornAnalysis
from OCR_Core_CustModel_HKT import runHKTAnalysis
from OCR_Core_ImageToPDF import convert2PDF
from OCR_Core_InvoiceFileHandler import splitPDF, getInvoiceIDList, mergeInvoiceFile, moveFile, renameFile, \
    removeStagingFiles
from OCR_Core_InvoiceFileHandler import checkFilePageCount
from OCR_Core_GetInvoiceData import getInvoiceData


def main():
    # Initialize Global Variables
    # --------------------------------------------------------
    config = ConfigParam()
    invoice_id = ''
    inputFolder = config.localInputFolderPath
    processFolder = config.localProcessFolderPath
    outputFolder = config.localOutputFolderPath
    archivedFolder = config.localArchiveFolderPath
    spPdfOutputFolder = config.targetInvoiceFileFolderPath
    spJsonOutputFolder = config.targetInvoiceJSONFolderPath
    resultSet = config.resultSet

    # Initialize Runtime Variables
    # --------------------------------------------------------
    # fileList = getInputFileList(inputFolder)  # file(s) lookup in Input Folder
    # fileListLen = len(fileList)
    processedPageCount = 0
    preCheckResult = False
    ocrMode = 1
    imageFile_Flag = False
    isHKT = False

    start_time = time.perf_counter()

    preCheckResult, fileListLen, fileList = preProcessChecker(inputFolder)

    if preCheckResult:
        print("OCR_Core_InvoiceRecognizer | Start Processing ......")
        for f in range(1, fileListLen):
            # set current file path
            inputFilePath = fileList.get(f)
            basename = os.path.basename(inputFilePath)

            # Special handling for HKT case
            if (basename[0:3]) == 'HKT':
                isHKT = True

            # Special handling for ChatHorn case
            if (basename[0:8]) == 'ChatHorn':
                isChatHorn = True

            if inputFilePath.lower().endswith('.jpg'):
                imageFile_name = os.path.splitext(basename)[0]
                imageFile_Flag = True
            else:
                file_name = os.path.splitext(basename)[0]

            # If input file is not an image
            if not imageFile_Flag:
                # Pre-Condition Check - Checking on raw file(s)
                # --------------------------------------------------------
                # @ 1 - Check original file page count
                pageCount = checkFilePageCount(inputFilePath)
                print("OCR_Core_InvoiceRecognizer | Check original file "
                      "\n > Filename : " + basename +
                      "\n > Page Count : " + str(pageCount))

                if pageCount == 1:
                    # --------------------------------------------------------
                    # @ 2 - Split PDF is not required if page number = 1
                    print("OCR_Core_InvoiceRecognizer | Single Page File Processing Initiated")

                    # 2.1) Get File Ready for Processing
                    # inputFilepath = inputFolder + rawFile
                    processFilePath = processFolder + basename
                    moveFile(inputFolder, processFolder, basename)
                    print(
                        "OCR_Core_InvoiceRecognizer | Single Page File Processing | Prepare Input file for Processing")

                    # 2.2) get Finalized JSON result object, and Invoice ID for naming
                    if isHKT:
                        # resultSet.append(runHKTAnalysis(processFilePath, processFilePath + '.json',
                        # 'application/pdf'))
                        resultSet.append(getInvoiceData(processFilePath, 1, ocrMode))  # with mode 1 to see details
                    elif isChatHorn:
                        resultSet.append(
                            runChatHornAnalysis(processFilePath, processFilePath + '.json', 'application/pdf'))
                    else:
                        resultSet.append(getInvoiceData(processFilePath, 1, ocrMode))  # with mode 1 to see details

                    resultSet, invoice_id = finalizedResultSet(resultSet, basename, 1)
                    lst = resultSet[0]
                    print("OCR_Core_InvoiceRecognizer | Single Page File Processing | Get Finalized JSON Result object")

                    # 2.3) Validate the result object
                    # print("Before validation : \n", lst)
                    validateExtractedRecord(lst, 'single')

                    # 2.4) Validate the result object
                    insertRecord(lst, 'single')

                    # 2.5) Generate JSON file to local output "2_Output"
                    outputJsonFile = writeJsonOutput(lst, outputFolder, file_name, 1)

                    print("OCR_Core_InvoiceRecognizer | Single Page File Processing | Generate JSON Output")

                    # 2.6) Rename Invoice File after process
                    outputPdfFileName = file_name + '_' + invoice_id + ".pdf"

                    newProcessFilePath = processFolder + outputPdfFileName
                    renameFile(processFilePath, newProcessFilePath)
                    print(
                        "OCR_Core_InvoiceRecognizer | Single Page File Processing | Rename Invoice File after process")

                    # 2.7) Move JSON file to SharePoint "22_Output" and Move File to SharePoint "21_PDF"
                    # moveFile(outputFolder, spJsonOutputFolder, outputJsonFile)
                    # moveFile(processFolder, spPdfOutputFolder, outputPdfFileName)
                    print("OCR_Core_InvoiceRecognizer | Single Page File Processing | Move File to SharePoint")

                    print("OCR_Core_InvoiceRecognizer | Single Page File Processing Completed!")
                    end_time = time.perf_counter()
                    elapsed_time = end_time - start_time

                else:
                    # --------------------------------------------------------
                    # @ 3 - Split PDF if page number > 1
                    print("OCR_Core_InvoiceRecognizer | Batch Processing Initiated")

                    # 3.1) Get File Ready for Processing
                    # inputFilepath = inputFolder + basename
                    inputFilepath = inputFolder + basename

                    # 3.2) File has more than 1 page, Page split is required and follow the procedure below
                    splitPDF(inputFilepath, processFolder)
                    print("OCR_Core_InvoiceRecognizer | Batch Processing | Splitting original files into staging")

                    # 3.3) Invoice ID Classification
                    splitFileList = getInputFileList(processFolder)
                    splitFileListLen = len(splitFileList)
                    # Invoke Form Recognizer to get Invoice ID for File Split and Merge
                    count = 0
                    for i in splitFileList:
                        if i != "id":
                            print("OCR_Core_InvoiceRecognizer | Batch Processing | Extracting data from staging file",
                                  str((count + 1)), "of", str((splitFileListLen - 1)))
                            resultSet.append(getInvoiceData(splitFileList[i], i, ocrMode))  # with mode 1 to see details
                            count = count + 1

                    # get Invoice Classifier Result for file merging process
                    tmpInvoiceID_List = []
                    tmp_resultStr = {}
                    tmpInvoiceID_List, tmp_resultStr = getInvoiceIDList(resultSet)
                    mergeInvoiceFile(tmpInvoiceID_List, tmp_resultStr, processFolder, outputFolder, splitFileList[1])
                    print("OCR_Core_InvoiceRecognizer | Batch Processing | Merging staging files")

                    # 3.4) Move Original File to Archived
                    moveFile(inputFolder, archivedFolder, basename)
                    print("OCR_Core_InvoiceRecognizer | Batch Processing | Moving the original file to archived folder")

                    # 3.5) Remove Staging Files
                    removeStagingFiles(tmpInvoiceID_List, tmp_resultStr, processFolder, splitFileList[1])
                    print("OCR_Core_InvoiceRecognizer | Batch Processing | Removing all staging files")

                    # 3.6) Invoke Form Recognizer to get final result set
                    finalResultSet = []
                    FileList = getInputFileList(outputFolder)
                    FileListLen = len(FileList)
                    # Invoke Form Recognizer to get Invoice ID for File Split and Merge
                    count = 0
                    for i in FileList:
                        if i != "id":
                            print("OCR_Core_InvoiceRecognizer | Batch Processing | Extracting data from merged file",
                                  str((count + 1)), "of", str((FileListLen - 1)))
                            finalResultSet.append(getInvoiceData(FileList[i], i, ocrMode))  # with mode 1 to see details
                            count = count + 1

                    # 3.7) get Finalized JSON result object
                    lst = finalizedResultSet(finalResultSet, None, 2)

                    # 3.8 Validate the result object
                    validateExtractedRecord(lst, 'multiple')

                    # 3.9 Validate the result object
                    insertRecord(lst, 'multiple')

                    # 3.10) Generate JSON result file and move to SharePoint
                    filename_list = getInputFileList(outputFolder)
                    filename_listLen = len(filename_list)

                    count = 0
                    for x in filename_list:
                        if x != "id":
                            basename = os.path.basename(filename_list[x])
                            file_name = os.path.splitext(basename)[0]
                            a = len(file_name)
                            b = len(lst[count]["InvoiceID"])
                            c = a - b
                            tmp_invoiceID = (file_name[c:c + b])
                            if tmp_invoiceID == lst[count]["InvoiceID"]:
                                # print(lst[count])
                                # print(file_name)
                                outputPdfFileName = file_name + ".pdf"
                                outputJsonFile = writeJsonOutput(lst[count], outputFolder, file_name, 2)
                                # moveFile(outputFolder, spJsonOutputFolder, outputJsonFile)
                                # moveFile(outputFolder, spPdfOutputFolder, outputPdfFileName)
                                print(
                                    "OCR_Core_InvoiceRecognizer | Batch Processing | Moving " + file_name + " to SharePoint")

                            count = count + 1

                    print("OCR_Core_InvoiceRecognizer | Batch Processing Completed!")
                    end_time = time.perf_counter()
                    elapsed_time = end_time - start_time

            # If input file is an image
            elif imageFile_Flag:

                # Sample Steps as Single File Processing
                # 2.1) Get File Ready for Processing
                # inputFilepath = inputFolder + rawFile
                processFilePath = processFolder + basename
                moveFile(inputFolder, processFolder, basename)
                print("OCR_Core_InvoiceRecognizer | Image File Processing | Prepare Input file for Processing")

                # 2.2) get Finalized JSON result object, and Invoice ID for naming
                resultSet.append(getInvoiceData(processFilePath, 1, ocrMode))  # with mode 1 to see details
                resultSet, invoice_id = finalizedResultSet(resultSet, basename, 1)
                lst = resultSet[0]
                print("OCR_Core_InvoiceRecognizer | Image File Processing | Get Finalized JSON Result object")

                # 2.3) Validate the result object
                print("Before validation : \n", lst)
                validateExtractedRecord(lst, 'single')

                # 2.4) Validate the result object
                insertRecord(lst, 'single')

                # 2.5) Generate JSON file to local output "2_Output"
                outputJsonFile = writeJsonOutput(lst, outputFolder, imageFile_name, 1)
                print("OCR_Core_InvoiceRecognizer | Image File Processing | Generate JSON Output")

                # 2.6) Rename Invoice File after process

                # 2.6) Rename Invoice File after process
                # Change JPG filename first
                outputJPGFileName = imageFile_name + '_' + invoice_id + ".jpg"
                newProcessJPGFilePath = processFolder + outputJPGFileName
                renameFile(processFilePath, newProcessJPGFilePath)
                print("OCR_Core_InvoiceRecognizer | Image File Processing | Rename Invoice File after process")

                # Define PDF filename
                outputPdfFileName = imageFile_name + '_' + invoice_id + ".pdf"
                newProcessPDFFilePath = processFolder + outputPdfFileName

                # Covert JPG to PDF
                convert2PDF(newProcessPDFFilePath, newProcessJPGFilePath)
                print(
                    "OCR_Core_InvoiceRecognizer | Image File Processing | Convert Invoice File from Image to PDF format")

                # 2.7) Move JSON file to SharePoint "22_Output" and Move File to SharePoint "21_PDF"
                # moveFile(outputFolder, spJsonOutputFolder, outputJsonFile)
                # moveFile(processFolder, spPdfOutputFolder, outputPdfFileName)
                # os.remove(newProcessJPGFilePath)
                print("OCR_Core_InvoiceRecognizer | Image File Processing | Move File to SharePoint")

                print("OCR_Core_InvoiceRecognizer | Image File Processing Completed!")
                end_time = time.perf_counter()
                elapsed_time = end_time - start_time

    print('\n-----------------------------------------------------\n' +
          'Total Elapsed Time : ' + str(datetime.timedelta(seconds=elapsed_time)) + '\n')

    if imageFile_Flag:
        print('An image file processed : ' + outputJPGFileName)
    elif pageCount >= 1:
        print('Total Number of File Processed : ' + str(pageCount))
