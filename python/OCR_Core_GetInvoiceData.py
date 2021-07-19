"""
OCR_Core_getInvoiceRecognized - Azure Form Recognizer functions

Sample:
AsiaPacificSecurity     [vendorName, BeneficiaryName, invoiceID, invoiceDate, poNum, invoiceTotal]
ChatHorn                [vendorName, BeneficiaryName, invoiceID, invoiceDate, poNum (Not detected),
                        invoiceTotal (use Amount Due if invoiceTotal is None)]
ChungShunLockSmith      [vendorName, BeneficiaryName, invoiceID, invoiceDate, poNum, invoiceTotal]
NewSky                  [vendorName, BeneficiaryName (use Vendor Name), invoiceID, invoiceDate, poNum, invoiceTotal]

CITIC (FIN_46011)
HKT


endpoint = "https://ocr-demo-01.cognitiveservices.azure.com"
credential = AzureKeyCredential("dfee77aeaed9445188878137903745b0")
"""
import os
from azure.ai.formrecognizer import FormRecognizerClient
from azure.core.credentials import AzureKeyCredential

from OCR_Config import ConfigParam

LogMsgHeader = os.path.basename(__file__) + " | "


# For Invoice Recognition
def getInvoiceData(PDFfilepath, pageID, mode):
    endpoint = ConfigParam.MSFTFormRecognizerEndpoint
    credential = AzureKeyCredential(ConfigParam.MSFTFormRecognizerApiKey)
    form_recognizer_client = FormRecognizerClient(endpoint, credential)

    # Create result set dictionary for capturing page record
    vendorName = ""
    BeneficiaryName = ""
    invoiceID = ""
    invoiceDate = ""
    poNum = ""
    invoiceTotal = ""
    customerID = ""

    with open(PDFfilepath, "rb") as f:
        poller = form_recognizer_client.begin_recognize_invoices(invoice=f, locale="en-US")
        invoices = poller.result()

    print(LogMsgHeader + getInvoiceData.__name__ + " Getting response for " + os.path.basename(PDFfilepath), "\n")
    for idx, invoice in enumerate(invoices):
        # print("--------Recognizing invoice #{}--------".format(idx + 1))
        if mode == 1: print("--------Recognizing invoice Page #", pageID, "----------------")

        # 1) Customer ID
        customer_id = invoice.fields.get("CustomerId")
        if customer_id:
            if mode == 1: print("Customer Id: {}".format(customer_id.value))
            customerID = ("{}".format(customer_id.value))

        # 2) Invoice Date
        invoice_date = invoice.fields.get("InvoiceDate")
        if invoice_date:
            if mode == 1: print("Invoice Date: {}".format(invoice_date.value))
            invoiceDate = str("{}".format(invoice_date.value))

        # 3) Invoice ID
        invoice_id = invoice.fields.get("InvoiceId")
        if invoice_id:
            if mode == 1: print("Invoice Id: {}".format(invoice_id.value))
            invoiceID = str("{}".format(invoice_id.value))

        # 4) Total Payment Amount
        invoice_total = invoice.fields.get("InvoiceTotal")
        if invoice_total:
            if mode == 1: print("Invoice Total: {}".format(invoice_total.value))
            invoiceTotal = str("{}".format(invoice_total.value))

        # 5) PO Number
        purchase_order = invoice.fields.get("PurchaseOrder")
        if purchase_order:
            if mode == 1: print("Purchase Order: {}".format(purchase_order.value))
            poNum = purchase_order.value

        # 6) Beneficiary Name / Remittance Address Recipient
        remittance_address_recipient = invoice.fields.get("RemittanceAddressRecipient")
        if remittance_address_recipient:
            if mode == 1: print("Remittance Address Recipient: {}".format(remittance_address_recipient.value))
            BeneficiaryName = str("{}".format(remittance_address_recipient.value))

        # 7) Company Name
        vendor_name = invoice.fields.get("VendorName")
        if vendor_name:
            if mode == 1: print("Vendor Name: {}".format(vendor_name.value))
            vendorName = ("{}".format(vendor_name.value))

        # Conditional handling for HKL
        # ---------------------------------------------------------------------------------------------

        # Chat Horn Engineering Ltd. - Invoice Total Amount
        amount_due = invoice.fields.get("AmountDue")
        if amount_due:
            if mode == 1: print("Amount Due: {}".format(amount_due.value))
            amountDue = str("{}".format(amount_due.value))

            if invoice_total == 'None':
                invoiceTotal = amountDue

        # HKT - Invoice Total Amount
        if (vendorName == 'HKT') or (vendorName == 'HKT Here- Serve HONG KONG LAND') or \
                (vendorName == 'THE HONGKONG LONGOGROUP'):
            vendorName == 'HKT'
            BeneficiaryName = vendorName
            invoiceID = customerID

        # Chung Shun - Beneficiary Name
        if vendorName == 'CHUNG SHUN LOCKSMITH':
            BeneficiaryName = vendorName

        # Asia Pacific Security - Vendor Name
        if 'APSS ASIA ' in vendorName:
            vendorName = 'ASIA PACIFIC SECURITY SERVICES LIMITED'

        # New Sky Construction Engineering Limited - Vendor Name and Beneficiary Name
        if ('WAS Construction' in vendorName) or ('WAS Sky' in vendorName) or (' Sky Construction' in vendorName):
            vendorName = 'New Sky Construction Engineering Limited'
            BeneficiaryName = vendorName

        # ---------------------------------------------------------------------------------------------

        if mode == 1: print("-----------------------------------------------------")

    pageResultSetStr = {'Page': pageID, 'CompanyName': vendorName, 'BeneficiaryName': BeneficiaryName,
                        'InvoiceID': invoiceID,
                        'InvoiceDate': invoiceDate, 'PONum': poNum, 'TotalAmount': invoiceTotal,
                        'CustomerId': customerID}

    return pageResultSetStr
