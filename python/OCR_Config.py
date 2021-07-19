# Core Configuration Parameters
# --------------------------------------------------------
class ConfigParam:

    # Form Recognizer Credentials
    MSFTFormRecognizerEndpoint = 'https://ocr-demo-01.cognitiveservices.azure.com'
    MSFTFormRecognizerApiKey = 'dfee77aeaed9445188878137903745b0'
    # ExploraFormRecognizerEndpoint = 'https://fr-hkl-demo.cognitiveservices.azure.com/'
    # ExploraFormRecognizerApiKey = '193affdd3a1f43958c60674a4421ec75'

    localRootFolder = "data/"
    localInputFolderPath = "data/0_Input/"
    localProcessFolderPath = "data/1_Process/"
    localOutputFolderPath = "data/2_Output/"
    localArchiveFolderPath = "data/3_Archived/"
    targetInvoiceFileFolderPath = "Z:\\Hong Kong Land\\99_POC\\2_Output\\21_PDF\\"
    targetInvoiceJSONFolderPath = "Z:\\Hong Kong Land\\99_POC\\2_Output\\22_JSON\\"

    # --------------------------------------------------------
    resultSet = [{'Page': '0', 'CompanyName': '', 'BeneficiaryName': '', 'InvoiceID': '',
                  'InvoiceDate': '', 'PONum': '', 'TotalAmount': '', 'CustomerId': ''}]

    # {"CustomerId": "", "InvoiceDate": "", "InvoiceId": "", "InvoiceTotal": "", "PurchaseOrder": "",
    # "RemittanceAddressRecipient": "", "VendorName": ""}

    dataFile = 'data/data.xlsx'
