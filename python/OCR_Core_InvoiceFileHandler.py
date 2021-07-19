import os
import shutil

import PyPDF2
from PyPDF2 import PdfFileWriter, PdfFileReader, PdfFileMerger
from OCR_Core_InvoiceClassifier import getInvoicePageNum, getUniqueInvoiceID

LogMsgHeader = os.path.basename(__file__) + " | "


def checkFilePageCount(PDFfilepath):
    with open(PDFfilepath, "rb") as f:
        pageCount = PdfFileReader(f).getNumPages()
        # print(LogMsgHeader + checkFilePageCount.__name__ + "")
    return pageCount


def moveFile(src_path, dst_path, file):
    f_src = os.path.join(src_path, file)
    if not os.path.exists(dst_path):
        os.mkdir(dst_path)
    f_dst = os.path.join(dst_path, file)
    shutil.move(f_src, f_dst)
    # print(LogMsgHeader + moveFile.__name__ + " | " + file + " has been moved for processing")


def renameFile(sourceFilepath, targetFilepath):
    shutil.move(sourceFilepath, targetFilepath)


def splitPDF(rawFilepath, processFolder):
    basename = os.path.basename(rawFilepath)
    file_name = os.path.splitext(basename)[0]
    # inputpdf = PdfFileReader(open(rawFilepath, "rb"))
    with open(rawFilepath, "rb") as f:
        inputpdf = PdfFileReader(f, "rb")

        for i in range(inputpdf.numPages):
            output = PdfFileWriter()
            output.addPage(inputpdf.getPage(i))
            page = i + 1
            with open(processFolder + file_name + "-page%s.pdf" % page, "wb") as outputStream:
                output.write(outputStream)
                filename = file_name + "-page%s.pdf" % page
                # print(LogMsgHeader + splitPDF.__name__ + " | Invoice File : " + filename + " is created")
                # print("OCR_Core_PdfHandler | splitPDF | Invoice File : " + filename + " is created")
            shutil.move(processFolder + file_name + "-page%s.pdf" % page, processFolder + file_name + "-page%s.pdf" % page)

    if i == page - 1:
        return True
    else:
        return False


def getInvoiceIDList(resultSet):
    # Perform File Merging
    # --------------------------------------------------------
    tmp_resultStr = {}
    tmp_invoiceID_list = []
    resultSet.pop(0)
    for a in resultSet:
        if a != '':
            invoice_id = a.get("InvoiceID")
            tmp_invoiceID_list.append(invoice_id)
            pageCount, resultStr, invoiceID_list = getInvoicePageNum(invoice_id, resultSet)
            # print(invoice_id, pageCount, resultStr)
            if pageCount == 1:
                for pageNum in invoiceID_list:
                    pos = pageNum.find("||", 0)
                    bPos = pos + 2
                    aPos = bPos - 2
                    pageStr = pageNum[0:aPos]
                    invoiceIdStr = pageNum[bPos:]
                    # print(pageStr, " | ", invoiceIdStr)
                    if invoiceIdStr == invoice_id:
                        resultStr = [pageStr]
                        # print(resultStr)
                        tmp_resultStr.update({invoice_id: resultStr})

            # page number > 1 will update the resultStr
            tmp_resultStr.update({invoice_id: resultStr})

    tmp_invoiceID_list = getUniqueInvoiceID(tmp_invoiceID_list)
    # tmp_invoiceID_list = tmp_invoiceID_list.tolist()
    # tmp_invoiceID_list.pop(0)
    # print("tmp_invoiceID_list | " + str(tmp_invoiceID_list))
    return tmp_invoiceID_list, tmp_resultStr


def mergeInvoiceFile(invoiceID_list, tmp_resultStr, processFolder, outputFolder, filePath):
    for i in invoiceID_list:
        # identify number of page detected per each invoice number (len >=1 can include all invoices)
        if len(tmp_resultStr[i]) >= 1:
            a = len(tmp_resultStr[i])
            page_list_str = tmp_resultStr[i]
            # print(page_list_str)

            # Call the PdfFileMerger
            # global file_name
            mergedObject = PdfFileMerger()
            pageListLen = len(page_list_str)
            basename = os.path.basename(filePath)

            file_name = os.path.splitext(basename)[0]
            pos = basename.find("-page")
            file_name = file_name[0:pos]

            # Loop through all of them and append their pages
            for fileNumber in range(0, pageListLen):
                # print("Page key : " + page_list_str[fileNumber])
                pageKey = page_list_str[fileNumber]
                # print("File info | " + inputFolder + file_name + "-page" + str(pageKey))
                mergedObject.append(
                    PdfFileReader(processFolder + file_name + "-page" + str(pageKey) + ".pdf", 'rb'))

            # Write all the files into a file which is named as shown below
            mergedObject.write(outputFolder + file_name + '_' + i + ".pdf")


def removeStagingFiles(invoiceID_list, tmp_resultStr, processFolder, filePath):
    for i in invoiceID_list:
        # identify number of page detected per each invoice number (len >=1 can include all invoices)
        if len(tmp_resultStr[i]) >= 1:
            a = len(tmp_resultStr[i])
            page_list_str = tmp_resultStr[i]
            # print(page_list_str)

            # Call the PdfFileMerger
            # global file_name
            mergedObject = PdfFileMerger()
            pageListLen = len(page_list_str)
            basename = os.path.basename(filePath)

            file_name = os.path.splitext(basename)[0]
            pos = basename.find("-page")
            file_name = file_name[0:pos]

            # Loop through all of them and append their pages
            for fileNumber in range(0, pageListLen):
                # print("Page key : " + page_list_str[fileNumber])
                pageKey = page_list_str[fileNumber]
                # print("File info | " + inputFolder + file_name + "-page" + str(pageKey))
                os.remove(processFolder + file_name + "-page" + str(pageKey) + ".pdf")
