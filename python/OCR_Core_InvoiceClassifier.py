import itertools
import numpy as np


# Testing Code
# invoiceID_list.append(str(key) + "||" + str(19955))
# invoiceID_list.append(str(key + 1) + "||" + str(19979))
# invoiceID_list.append(str(key + 2) + "||" + str(19955))

def getPageList(alist):
    x = np.array(alist)
    x = np.unique(x)
    # Convert Numpy Object to List
    return x.tolist()


def getUniqueInvoiceID(alist):
    x = np.array(alist)
    return np.unique(x)


def getInvoicePageNum(invoiceID, resultSet):
    invoiceID_list = []
    pageNumList = []
    counter = 0
    key = 0

    # Concatenate Page number and Invoice ID in [{Page Number}||{Invoice ID}]
    for x in resultSet:
        # print(x)
        page_id = x["Page"]
        invoice_id = x["InvoiceID"]
        invoiceID_list.append(str(page_id) + "||" + str(invoice_id))
        key = key + 1
    # invoiceID_list.remove('0||')
    # print(invoiceID_list)

    loopCount = 0
    for a, b in itertools.combinations(invoiceID_list, 2):
        aPos = a.find("||", 0)
        bPos = b.find("||", 0)
        aPos = aPos + 2
        bPos = bPos + 2
        aStr = a[aPos:]
        bStr = b[bPos:]
        # print(aStr, bStr)
        if aStr == bStr:
            aPos = aPos - 2
            bPos = bPos - 2
            aPageStr = a[0:aPos]
            bPageStr = b[0:bPos]
            # print("Matched Invoice ID #", aStr, "found on Page #", aPageStr + " Page #", bPageStr)
            if aStr == invoiceID:
                pageNumList.append(aPageStr)
                pageNumList.append(bPageStr)
                counter = counter + 1

    # print("Invoice ID #", invoiceID, 'appeared for ', counter, 'times')
    # print(pageNum)
    counter = counter + 1
    return counter, getPageList(pageNumList), invoiceID_list
