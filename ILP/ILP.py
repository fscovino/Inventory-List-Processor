'''
Created on November 01, 2016
ILP - Inventory List Processor v1.0
Loads an Excel file with a list of inventory items to refine and customize columns into a format ready to share with customers.
@author: Francisco Scovino - CMH United Corp (Miami, Fl)
'''
# List of Imports
import tkinter
from tkinter import ttk
from tkinter import filedialog
import xml.etree.ElementTree as xmlet
import xlsxwriter
import datetime
import csv
import smtplib
import os.path
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from _io import open
import win32com.client as win32

# Constant Values
IMG_DOMAIN = 'http://www.cmhunited.com/images/'
IMG_EXTENSION = '.jpg'
IMG_FOLDER = 'Z:/Workstations/Francisco/Art/cmhunited.com/images/'
MAIL_SMTP = 'smtp.gmail.com:587'
MAIL_USR = 'franciscocmh'
MAIL_PWD = 'barefoot6561'
MAIL_FROM = 'francisco@cmhcellular.com'
MAIL_SUBJECT = 'Inventory List'
MAIL_SIGNATURE_IMG = 'Z:/Public/Signatures/Francisco_Signature.png'

# Excel Format Dictionaries and Columns Width
dicHeaderFormat = {'bold':True, 'bg_color':'black', 'font_color':'white', 'bottom':2, 'bottom_color':'red', 'font_size':9, 'align':'center'}
dicBrandFormat = {'bold': True, 'bg_color':'black', 'font_color':'white', 'right':2, 'border_color':'red', 'font_size':9, 'align':'center'}
dicCornerFormat = {'bold':True, 'bg_color':'red', 'font_color':'black', 'bottom':2, 'bottom_color':'red', 'font_size':9, 'align':'center'}
dicRateFormat = {'font_size':9, 'align':'center', 'num_format':'0.00'}
dicCenterFormat = {'font_size':9, 'align':'center'}
dicLeftFormat = {'font_size':9, 'align':'left'}
dicUrlFormat = {'font_size':9, 'align':'center', 'underline':1, 'font_color':'blue'}

# Column Width
colWidthBrand = 15
colWidthPart = 30
colWidthDesc = 60

# Logic Methods
def loadAndProcess(filePath, isQBgo):
    '''Load .xls file from directory and populate an array rawList. Process each field through different filters to produce a clean list.
    0:Brand, 1:Part#, 2:Description, 3:Available, 4:Coming, 5:Image, 6:Level0, 7:Level1, 8:Level2, 9:Level3, 10:Barcode
    '''
    cleanList = []
    # Dictionary for select case on mapping columns
    caseColumn = {'Column A':0, 'Column B':1, 'Column C':2, 'Column D':3, 'Column E':4, 'Column F':5, 'Column G':6, 'Column H':7}
    # Mapping Columns
    mPart = caseColumn[varPartNumber.get()]
    mDesc = caseColumn[varDescription.get()]
    mCost = caseColumn[varCost.get()]
    mPrice = caseColumn[varPrice.get()]
    mQoH = caseColumn[varQtyHand.get()]
    mQoS = caseColumn[varQtySO.get()]
    mQoP = caseColumn[varQtyPO.get()]
    mMpn = caseColumn[varMpn.get()]
    
    # Populate cleanlist with data from QuickBooks
    if isQBgo == 1:
        # Export List Directly from QuickBooks
        qbList = importQbInventory()
        # 0:Part, 1:Desc, 2:qtyOH, 3:qtySO, 4:qtyPO, 5:cost, 6:price, 7:mpn
        for row in qbList:
            # Avoid items with NA on brand
            if filterGetBrand(row[1]) == 'NA':
                continue
            item = []
            item.append(filterGetBrand(row[1]))                                # Brand
            item.append(row[0])                                                # Part
            item.append(filterCleanDescription(row[1]))                        # Description
            item.append(filterGetAvailable(row[2], row[3]))                    # Available
            item.append(filterGetInteger(row[4]))                              # Coming
            item.append(filterGetImageLink(row[0]))                            # Image
            item.append(0.00)                                                  # Level 0
            item.append(filterGetFloat(row[5]))                                # Level 1
            item.append(filterGetFloat(row[6]))                                # Level 2
            item.append(filterAddPercentage(row[6], 5))                        # Level 3
            item.append(row[7])                                                # Barcode
            # Add the item to te clean list
            cleanList.append(item)
    else:
        # Populate cleanlist with external data
        try:
            # Open the .csv file with rawData
            with open(filePath, 'r') as rawFile:
                # Read file into an array
                rawList = csv.reader(rawFile)
                # Iterate through each row
                for raw in rawList:
                    # Avoid items with NA on brand
                    if filterGetBrand(raw[mDesc]) == 'NA':
                        continue
                    item = []
                    item.append(filterGetBrand(raw[mDesc]))                    # Brand
                    item.append(raw[mPart])                                    # Part
                    item.append(filterCleanDescription(raw[mDesc]))            # Description
                    item.append(filterGetAvailable(raw[mQoH], raw[mQoS]))      # Available
                    item.append(filterGetInteger(raw[mQoP]))                   # Coming
                    item.append(filterGetImageLink(raw[mPart]))                # Image
                    item.append(0.00)                                          # Level 0
                    item.append(filterGetFloat((raw[mCost])))                  # Level 1
                    item.append(filterGetFloat(raw[mPrice]))                   # Level 2
                    item.append(filterAddPercentage(raw[mPrice], 5))           # Level 3
                    item.append(raw[mMpn])                                     # Barcode
                    # Add the item to te clean list
                    cleanList.append(item)
        
        except Exception as inst:
            print('Error on: LoadAndProcess:')
            print(inst)
            
    # Sort list for Brand then for Description
    cleanList.sort(key=lambda x: (x[0], x[2]))
    print('*** File Loaded and Processed Succesfully')
    return cleanList
            
def filterCleanDescription(description):
    '''Clean description by replacing characters and trimming spaces'''
    val1 = description.replace(',', ' /')
    val2 = val1.replace('*', 'Apple')
    val3 = filterTrimSpaces(val2)
    return val3

def filterTrimSpaces(description):
    '''Trim all extra spaces to single spaces'''
    val1 = description.replace(' ', '{}')
    val2 = val1.replace('}{', '')
    val3 = val2.replace('{}', ' ')
    return val3

def filterGetBrand(description):
    '''Extract the brand from the description field'''
    brand = 'NA'
    # Separate line into words on an array
    words = description.replace('*', 'Apple').split()
    # Select the second word if first word = original
    if words[0].lower() == 'original':
        brand = words[1].upper()

    # Complete brands that have two words
    if brand == 'AGENT':
        brand += ' 18'
    elif brand == 'BODY':
        brand += ' GLOVE'
    elif brand == 'BLACK':
        brand += ' ROCK'
    elif brand == 'PURE':
        brand += 'GEAR'
    elif brand == 'WHITE':
        brand += ' DIAMONDS'

    return brand

def filterGetAvailable(qtyOnHand, qtyOnSales):
    '''Retrieve the difference between 2 quantities''' 
    val1 = filterGetInteger(qtyOnHand)
    
    val2 = filterGetInteger(qtyOnSales)
        
    result = val1 - val2
    
    return result

def filterGetInteger(num):
    '''Verify a string and turn it into a valid integed'''
    val1 = num.replace(',', '').split('.')[0]
    if val1.isnumeric():
        val1 = int(val1)
    else:
        val1 = 0
        
    return val1

def filterGetFloat(num):
    '''Verify a string and turn it into a valid float'''
    val1 = num.replace(',', '')
    if val1 == '' or val1 == '0' or val1.isalpha():
        val1 = 0.00
    else:
        try:
            val1 = float(val1)
        except ValueError:
            print('Error on filterGetFloat:')
                    
    return val1

def filterGetImageLink(partNumber):
    '''Build and retrieve a hyperlink based on part number given'''
    value = IMG_DOMAIN + partNumber + IMG_EXTENSION
    return value

def filterAddPercentage(amount, percentage):
    '''Add a percentage to a given value'''
    val1 = filterGetFloat(amount)
        
    result = val1 + ((val1 * percentage) / 100)
    
    return round(result, 1)

def importQbInventory():
    '''Connect to QuickBooks and extract Inventory List'''
    APP_CODE = ''
    APP_NAME = 'ILP - Inventory List'
    QB_FILE = ''
    QB_OPEN_MODE_DO_NOT_CARE = 2
    
    qb = None
    ticket = ''
    response = ''
    inventory = []
    
    try:
        # Get QuickBooks Object
        qb = win32.Dispatch('QBXMLRP2.RequestProcessor')
        # Open Connection
        qb.OpenConnection(APP_CODE, APP_NAME)
        # Begin Session
        ticket = qb.BeginSession(QB_FILE, QB_OPEN_MODE_DO_NOT_CARE)
        if ticket != '':
            print('*** Connected to QuickBooks Succesfully')
        
    except Exception as inst:
        print('Error on importQbInventory')
        print(inst)
        
    else:
        # Build Prolog
        prolog = '<?xml version=\"1.0\" encoding=\"utf-8\"?><?qbxml version=\"7.0\"?>'
        # Build xml Request
        qbxml = xmlet.Element('QBXML')
        qbxmlMsgsRq = xmlet.SubElement(qbxml, 'QBXMLMsgsRq', onError='stopOnError')
        itemInventoryQueryRq = xmlet.SubElement(qbxmlMsgsRq, 'ItemInventoryQueryRq')
        str(itemInventoryQueryRq)
        # Assemble Request
        xmlRequest = prolog + xmlet.tostring(qbxml).decode('utf-8')
        # Process Request
        response = qb.ProcessRequest(ticket, xmlRequest)
        # Process Response
        root = xmlet.fromstring(response)
        msgsRs = root.find('QBXMLMsgsRs')
        # Make sure we get a response
        if len(msgsRs) == 1:
            # Get the status value
            status = msgsRs[0].get('statusCode')
            # Read nodes if status is OK = 0
            if int(status) == 0:
                # Get a list of items
                itemList = msgsRs.find('ItemInventoryQueryRs')
                # Add each item to the inventory list
                for item in itemList:
                    # 0:Part, 1:Desc, 2:qtyOH, 3:qtySO, 4:qtyPO, 5:cost, 6:price, 7:mpn
                    row = ['', '', '', '', '', '', '', '']
                    row[0] = item.find('FullName').text
                    # Verify Description Node Exist
                    if item.find('SalesDesc') != None:
                        row[1] = item.find('SalesDesc').text
                    else:
                        row[1] = 'NA'
                    row[2] = item.find('QuantityOnHand').text
                    row[3] = item.find('QuantityOnSalesOrder').text
                    row[4] = item.find('QuantityOnOrder').text
                    row[5] = item.find('PurchaseCost').text
                    row[6] = item.find('SalesPrice').text
                    # Verify MPN Node Exist
                    if item.find('ManufacturerPartNumber') != None:
                        row[7] = item.find('ManufacturerPartNumber').text
                        
                    # Add item to the inventory
                    inventory.append(row)
                    
    finally:
        # End Session
        if ticket != '':
            qb.EndSession(ticket)
            
        # Close Connection           
        if qb != None:
            qb.CloseConnection()
            # Drop QuickBooks Oblect
            qb = None
            
        # Return the inventory list
        return inventory

def exportListNoPrices(cleanList, filePath):
    '''Export to a .xls file an array list to the given path.
    This list has no prices.
    '''
    # Get today's date for the file name
    now = str(datetime.datetime.now()).split()[0]
    # Create a workbook on the specified folder
    fileName = '/Inv_' + now + '.xlsx'
    wb = xlsxwriter.Workbook(filePath + fileName)
    # Create Formats
    fmtHeader = wb.add_format(dicHeaderFormat)
    fmtBrand = wb.add_format(dicBrandFormat)
    fmtCorner = wb.add_format(dicCornerFormat)
    fmtCenter = wb.add_format(dicCenterFormat)
    fmtUrl = wb.add_format(dicUrlFormat)
    fmtLeft = wb.add_format(dicLeftFormat)
    # Create a Worksheet
    ws = wb.add_worksheet()
    # Create the header
    ws.write_string('A1', 'Brand', fmtCorner)
    ws.write_string('B1', 'Part #', fmtHeader)
    ws.write_string('C1', 'Description', fmtHeader)
    ws.write_string('D1', 'Available', fmtHeader)
    ws.write_string('E1', 'Coming', fmtHeader)
    ws.write_string('F1', 'Image', fmtHeader)
    # Start passing data from the array to the xls file
    row, col = 1, 0
    for item in cleanList: #arrInventory
        ws.write_string(row, col + 0, item[0], fmtBrand)    # Brand
        ws.write_string(row, col + 1, item[1], fmtLeft)     # Part
        ws.write_string(row, col + 2, item[2], fmtLeft)     # Description
        ws.write_number(row, col + 3, item[3], fmtCenter)   # Available
        ws.write_number(row, col + 4, item[4], fmtCenter)   # Coming
        ws.write_url(row, col + 5, item[5], fmtUrl, 'Image')# Image
        # Increase row number
        row += 1
        
    # Set Column Width
    ws.set_column('A:A', colWidthBrand)
    ws.set_column('B:B', colWidthPart)
    ws.set_column('C:C', colWidthDesc)
    # Close Workbook when finished
    wb.close()
    # Return the path of the new created file to be sent by email
    print('*** List No Prices Exported Succesfully')
    return filePath + fileName

def exportListHighPrices(cleanList, filePath):
    '''Export to a .xls file an array list to the given path.
    This list has only the high prices (price from QB).
    '''
    # Get today's date for the file name
    now = str(datetime.datetime.now()).split()[0]
    # Create a workbook on the specified folder
    fileName = '/Inv_' + now + '_wp' + '.xlsx'
    wb = xlsxwriter.Workbook(filePath + fileName)
    # Create Formats
    fmtHeader = wb.add_format(dicHeaderFormat)
    fmtBrand = wb.add_format(dicBrandFormat)
    fmtCorner = wb.add_format(dicCornerFormat)
    fmtRate = wb.add_format(dicRateFormat)
    fmtCenter = wb.add_format(dicCenterFormat)
    fmtUrl = wb.add_format(dicUrlFormat)
    fmtLeft = wb.add_format(dicLeftFormat)
    # Create a Worksheet
    ws = wb.add_worksheet()
    # Create the header
    ws.write_string('A1', 'Brand', fmtCorner)
    ws.write_string('B1', 'Part #', fmtHeader)
    ws.write_string('C1', 'Description', fmtHeader)
    ws.write_string('D1', 'Available', fmtHeader)
    ws.write_string('E1', 'Coming', fmtHeader)
    ws.write_string('F1', 'Image', fmtHeader)
    ws.write_string('G1', 'Price', fmtHeader)

    # Start passing data from the array to the xls file
    row, col = 1, 0
    for item in cleanList: #arrInventory
        ws.write_string(row, col + 0, item[0], fmtBrand)    # Brand
        ws.write_string(row, col + 1, item[1], fmtLeft)     # Part
        ws.write_string(row, col + 2, item[2], fmtLeft)     # Description
        ws.write_number(row, col + 3, item[3], fmtCenter)   # Available
        ws.write_number(row, col + 4, item[4], fmtCenter)   # Coming
        ws.write_url(row, col + 5, item[5], fmtUrl, 'Image')# Image
        ws.write_number(row, col + 6, item[8], fmtRate)     # Level 2
        # Increase row number
        row += 1
        
    # Set Column Width
    ws.set_column('A:A', colWidthBrand)
    ws.set_column('B:B', colWidthPart)
    ws.set_column('C:C', colWidthDesc)
    # Close Workbook when finished
    wb.close()
    # Return the path of the new created file to be sent by email
    print('*** List High Prices Exported Succesfully')
    return filePath + fileName

def exportListTwoPrices(cleanList, filePath):
    '''Export to a .xls file an array list to the given path.
    This list has the low and high prices (cost and price from QB).
    '''
    # Get today's date for the file name
    now = str(datetime.datetime.now()).split()[0]
    # Create a workbook on the specified folder
    fileName = '/Inv_' + now + '_w2p' + '.xlsx'
    wb = xlsxwriter.Workbook(filePath + fileName)
    # Create Formats
    fmtHeader = wb.add_format(dicHeaderFormat)
    fmtBrand = wb.add_format(dicBrandFormat)
    fmtCorner = wb.add_format(dicCornerFormat)
    fmtRate = wb.add_format(dicRateFormat)
    fmtCenter = wb.add_format(dicCenterFormat)
    fmtUrl = wb.add_format(dicUrlFormat)
    fmtLeft = wb.add_format(dicLeftFormat)
    # Create a Worksheet
    ws = wb.add_worksheet()
    # Create the header
    ws.write_string('A1', 'Brand', fmtCorner)
    ws.write_string('B1', 'Part #', fmtHeader)
    ws.write_string('C1', 'Description', fmtHeader)
    ws.write_string('D1', 'Available', fmtHeader)
    ws.write_string('E1', 'Coming', fmtHeader)
    ws.write_string('F1', 'Image', fmtHeader)
    ws.write_string('G1', 'Cost', fmtHeader)
    ws.write_string('H1', 'Price', fmtHeader)
    # Start passing data from the array to the xls file
    row, col = 1, 0
    for item in cleanList: #arrInventory
        ws.write_string(row, col + 0, item[0], fmtBrand)    # Brand
        ws.write_string(row, col + 1, item[1], fmtLeft)     # Part
        ws.write_string(row, col + 2, item[2], fmtLeft)     # Description
        ws.write_number(row, col + 3, item[3], fmtCenter)   # Available
        ws.write_number(row, col + 4, item[4], fmtCenter)   # Coming
        ws.write_url(row, col + 5, item[5], fmtUrl, 'Image')# Image
        ws.write_number(row, col + 6, item[7], fmtRate)     # Level 1
        ws.write_number(row, col + 7, item[8], fmtRate)     # Level 2
        # Increase row number
        row += 1
        
    # Set Column Width
    ws.set_column('A:A', colWidthBrand)
    ws.set_column('B:B', colWidthPart)
    ws.set_column('C:C', colWidthDesc)
    # Close Workbook when finished
    wb.close()
    # Return the path of the new created file to be sent by email
    print('*** List Two Prices Exported Succesfully')
    return filePath + fileName

def exportListUpcPrices(cleanList, filePath):
    '''Export to a .xls file an array list to the given path.
    This list has only the high prices and Barcode (price and MPN from QB).
    '''
    # Get today's date for the file name
    now = str(datetime.datetime.now()).split()[0]
    # Create a workbook on the specified folder
    fileName = '/Inv_' + now + '_wp+upc' + '.xlsx'
    wb = xlsxwriter.Workbook(filePath + fileName)
    # Create Formats
    fmtHeader = wb.add_format(dicHeaderFormat)
    fmtBrand = wb.add_format(dicBrandFormat)
    fmtCorner = wb.add_format(dicCornerFormat)
    fmtRate = wb.add_format(dicRateFormat)
    fmtCenter = wb.add_format(dicCenterFormat)
    fmtUrl = wb.add_format(dicUrlFormat)
    fmtLeft = wb.add_format(dicLeftFormat)
    # Create a Worksheet
    ws = wb.add_worksheet()
    # Create the header
    ws.write_string('A1', 'Brand', fmtCorner)
    ws.write_string('B1', 'Part #', fmtHeader)
    ws.write_string('C1', 'Description', fmtHeader)
    ws.write_string('D1', 'Available', fmtHeader)
    ws.write_string('E1', 'Coming', fmtHeader)
    ws.write_string('F1', 'Image', fmtHeader)
    ws.write_string('G1', 'Price', fmtHeader)
    ws.write_string('H1', 'Barcode', fmtHeader)
    # Start passing data from the array to the xls file
    row, col = 1, 0
    for item in cleanList: #arrInventory
        ws.write_string(row, col + 0, item[0], fmtBrand)    # Brand
        ws.write_string(row, col + 1, item[1], fmtLeft)     # Part
        ws.write_string(row, col + 2, item[2], fmtLeft)     # Description
        ws.write_number(row, col + 3, item[3], fmtCenter)   # Available
        ws.write_number(row, col + 4, item[4], fmtCenter)   # Coming
        ws.write_url(row, col + 5, item[5], fmtUrl, 'Image')# Image
        ws.write_number(row, col + 6, item[8], fmtRate)     # Level 2
        ws.write_string(row, col + 7, item[10], fmtCenter)  # Barcode
        # Increase row number
        row += 1
        
    # Set Column Width
    ws.set_column('A:A', colWidthBrand)
    ws.set_column('B:B', colWidthPart)
    ws.set_column('C:C', colWidthDesc)
    # Close Workbook when finished
    wb.close()
    # Return the path of the new created file to be sent by email
    print('*** List UPC Prices Exported Succesfully')
    return filePath + fileName

def exportListWebPrices(cleanList, filePath):
    '''Export to a .csv file an array list to the given path.
    This list has all level of prices (cost, price and +5%).
    '''
    # Get today's date for the file name
    now = str(datetime.datetime.now()).split()[0]
    # Create a file on the specified folder
    fileName = filePath + '/Inv_' + now + '_web.csv'
    # Create a web list out of the cleanlist
    webList = [['Brand', 'Part', 'Description', 'Available', 'Coming', 'Level 1', 'level 2', 'Level 3']]
    # Create a row with custom values
    for item in cleanList:
        row = []
        row.append(item[0])
        row.append(item[1])
        row.append(item[2])
        row.append(item[3])
        row.append(item[4])
        row.append(item[7])
        row.append(item[8])
        row.append(item[9])
        # Append row into weblist
        webList.append(row)
        
    try:
        # Create a csv writer
        writer = csv.writer(open(fileName, 'w', newline=''))
        for item in webList:
            writer.writerow(item)
            
    except Exception as inst:
        print('Error on: exportListWebPrices:')
        print(inst)
    else:
        print('*** List Web Prices Exported Succesfully')

def exportMissingPics(cleanList, filePath):
    '''Export to a .xls file all missing images listed on the current inventory'''
    # Get today's date for the file name
    now = str(datetime.datetime.now()).split()[0]
    # Create a workbook on the specified folder
    wb = xlsxwriter.Workbook(filePath + '/MissingPics_' + now + '.xlsx')
    # Create a Worksheet
    ws = wb.add_worksheet()
    # Add part# that are not in the image folder to the list
    row = 0
    for item in cleanList:
        if os.path.isfile(IMG_FOLDER + item[1] + IMG_EXTENSION):
            continue
        else:
            ws.write_string(row, 0, item[1])
            row += 1
    # Close Workbook when finished
    wb.close()
    print('*** List Missing Images Exported Succesfully')

# Form Methods
def browseFile():
    varFileLocation.set(filedialog.askopenfilename())

def browseFolder():
    varExportFolder.set(filedialog.askdirectory())

def modifyEmailList(alist, operation):
    '''Add or Remove email addresses to the different email lists'''
    # Modify List No Prices
    if alist == 'lnp':
        if operation == 'add' and varListNoPrices.get() not in arrListNoPrices:
            arrListNoPrices.append(varListNoPrices.get())
        elif operation == 'delete' and varListNoPrices.get() != 'Subscribe' and varListNoPrices.get() in arrListNoPrices:
            arrListNoPrices.remove(varListNoPrices.get())
        # Update Widget with new data
        cboNoPrices['values'] = arrListNoPrices
        varListNoPrices.set(arrListNoPrices[0])

    # Modify List High Prices
    if alist == 'lhp':
        if operation == 'add' and varListHighPrices.get() not in arrListHighPrices:
            arrListHighPrices.append(varListHighPrices.get())
        elif operation == 'delete' and varListHighPrices.get() != 'Subscribe' and varListHighPrices.get() in arrListHighPrices:
            arrListHighPrices.remove(varListHighPrices.get())
        # Update Widget with new data
        cboHighPrices['values'] = arrListHighPrices
        varListHighPrices.set(arrListHighPrices[0])

    # Modify List Two Prices
    if alist == 'l2p':
        if operation == 'add' and varListTwoPrices.get() not in arrListTwoPrices:
            arrListTwoPrices.append(varListTwoPrices.get())
        elif operation == 'delete' and varListTwoPrices.get() != 'Subscribe' and varListTwoPrices.get() in arrListTwoPrices:
            arrListTwoPrices.remove(varListTwoPrices.get())
        # Update Widget with new data
        cboTwoPrices['values'] = arrListTwoPrices
        varListTwoPrices.set(arrListTwoPrices[0])

    # Modify List Upc Prices
    if alist == 'lup':
        if operation == 'add' and varListUpcPrices.get() not in arrListUpcPrices:
            arrListUpcPrices.append(varListUpcPrices.get())
        elif operation == 'delete' and varListUpcPrices.get() != 'Subscribe' and varListUpcPrices.get() in arrListUpcPrices:
            arrListUpcPrices.remove(varListUpcPrices.get())
        # Update Widget with new data
        cboUpcPrices['values'] = arrListUpcPrices
        varListUpcPrices.set(arrListUpcPrices[0])

    # Modify List Web Prices
    if alist == 'lwp':
        if operation == 'add' and varListWebPrices.get() not in arrListWebPrices:
            arrListWebPrices.append(varListWebPrices.get())
        elif operation == 'delete' and varListWebPrices.get() != 'Subscribe' and varListWebPrices.get() in arrListWebPrices:
            arrListWebPrices.remove(varListWebPrices.get())
        # Update Widget with new data
        cboWebPrices['values'] = arrListWebPrices
        varListWebPrices.set(arrListWebPrices[0])

def exportLists():
    '''Export lists and get the location of each price list to be sent by email'''
    # 0:No Prices, 1:High Prices, 2:Two Prices, 3:Upc Prices, 4:Web Prices
    lists = [0, 0, 0, 0, 0]
    # Load and Process the raw list from file
    mainList = loadAndProcess(varFileLocation.get(), varIsQbgo.get())
    
    if len(mainList) > 1:
        
        # Export List No Prices and get file location
        if varChkNoPrices.get() == 1:
            lists[0] = exportListNoPrices(mainList, varExportFolder.get())
        
        # Export List High Prices and get file location
        if varChkHighPrices.get() == 1:
            lists[1] = exportListHighPrices(mainList, varExportFolder.get())
        
        # Export List Two Prices and get file location
        if varChkTwoPrices.get() == 1:
            lists[2] = exportListTwoPrices(mainList, varExportFolder.get())
        
        # Export List UPC Prices and get file location
        if varChkUpcPrices.get() == 1:
            lists[3] = exportListUpcPrices(mainList, varExportFolder.get())

        # Export List Web Prices and get file location
        if varChkWebPrices.get() == 1:
            lists[4] = exportListWebPrices(mainList, varExportFolder.get())
            
        # Find and export a list with part numbers of the missing images
        if varExportImages.get() == 1:
            exportMissingPics(mainList, varExportFolder.get())
    
    # Return path of all lists to be send by email
    return lists
    
def sendListsViaGmail():
    '''List of paths with inventory to send'''
    # 0:No Prices, 1:High Prices, 2:Two Prices, 3:Upc Prices, 4:Web Prices
    # Export all lists ang get the ones that need to be emailed
    elist = exportLists()
    counter = 0
    
    try:
        # Send message to each selected list
        # Send List No Prices
        if elist[0] != 0:
            # Build message
            msg = MIMEMultipart()
            msg['From'] = MAIL_FROM
            msg['Reply-to'] = MAIL_FROM
            msg['Subject'] = MAIL_SUBJECT
            
            part = MIMEBase('application', "octet-stream")
            part.set_payload(open(elist[0],"rb").read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(elist[0]))
            msg.attach(part)
            # Refine list of subscribers. Ommit 'Subscribe" from list
            for subscriber in arrListNoPrices:
                if subscriber == 'Subscribe':
                    continue
                else:
                    recipients = subscriber.strip().split(',')
                    
            # Build Server
            smtp = smtplib.SMTP(MAIL_SMTP)
            # Start Server
            smtp.ehlo()
            smtp.starttls()
            smtp.login(MAIL_USR, MAIL_PWD)
            # Send the message
            smtp.sendmail(msg['From'], recipients, msg.as_string())
            print('*** List No Prices Sent Succesfully')
            counter += 1
        
        # Send List High Prices
        if elist[1] != 0:
            # Build message
            msg = MIMEMultipart()
            msg['From'] = MAIL_FROM
            msg['Reply-to'] = MAIL_FROM
            msg['Subject'] = MAIL_SUBJECT + ' wp'
            
            part = MIMEBase('application', "octet-stream")
            part.set_payload(open(elist[1],"rb").read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(elist[1]))
            msg.attach(part)
            # Refine list of subscribers. Ommit 'Subscribe" from list
            for subscriber in arrListHighPrices:
                if subscriber == 'Subscribe':
                    continue
                else:
                    recipients = subscriber.strip().split(',')
            # Build Server
            smtp = smtplib.SMTP(MAIL_SMTP)
            # Start Server
            smtp.ehlo()
            smtp.starttls()
            smtp.login(MAIL_USR, MAIL_PWD)
            # Send the message
            smtp.sendmail(msg['From'], recipients, msg.as_string())
            print('*** List High Prices Sent Succesfully')
            counter += 1
            
        # Send List Two Prices
        if elist[2] != 0:
            # Build message
            msg = MIMEMultipart()
            msg['From'] = MAIL_FROM
            msg['Reply-to'] = MAIL_FROM
            msg['Subject'] = MAIL_SUBJECT + ' wp'
            
            part = MIMEBase('application', "octet-stream")
            part.set_payload(open(elist[2],"rb").read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(elist[2]))
            msg.attach(part)
            # Refine list of subscribers. Ommit 'Subscribe" from list
            for subscriber in arrListTwoPrices:
                if subscriber == 'Subscribe':
                    continue
                else:
                    recipients = subscriber.strip().split(',')
            # Build Server
            smtp = smtplib.SMTP(MAIL_SMTP)
            # Start Server
            smtp.ehlo()
            smtp.starttls()
            smtp.login(MAIL_USR, MAIL_PWD)
            # Send the message
            smtp.sendmail(msg['From'], recipients, msg.as_string())
            print('*** List Two Prices Sent Succesfully')
            counter += 1
            
        # Send List Upc Prices
        if elist[3] != 0:
            # Build message
            msg = MIMEMultipart()
            msg['From'] = MAIL_FROM
            msg['Reply-to'] = MAIL_FROM
            msg['Subject'] = MAIL_SUBJECT + ' wp'
            
            part = MIMEBase('application', "octet-stream")
            part.set_payload(open(elist[3],"rb").read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(elist[3]))
            msg.attach(part)
            # Refine list of subscribers. Ommit 'Subscribe" from list
            for subscriber in arrListUpcPrices:
                if subscriber == 'Subscribe':
                    continue
                else:
                    recipients = subscriber.strip().split(',')
            # Build Server
            smtp = smtplib.SMTP(MAIL_SMTP)
            # Start Server
            smtp.ehlo()
            smtp.starttls()
            smtp.login(MAIL_USR, MAIL_PWD)
            # Send the message
            smtp.sendmail(msg['From'], recipients, msg.as_string())
            print('*** List UPC Prices Sent Succesfully')
            counter += 1
            
    except Exception as inst:
        print('Error on: sendListsViaGmail:')
        print(inst)
        
    else:
        print('*** ' + str(counter) + ' Selected List(s) were succesfully sent !!!')
        
    finally:
        # Close Server Connection
        smtp.quit()
        
def sendListsViaOutlook():
    '''List of paths with inventory to send'''
    # 0:No Prices, 1:High Prices, 2:Two Prices, 3:Upc Prices, 4:Web Prices
    # Export all lists ang get the ones that need to be emailed
    elist = exportLists()
    counter = 0
    
    try:
        # Get Outlook Object
        outlook = win32.Dispatch('Outlook.Application')
        
        # Send message to each selected list
        # Send List No Prices
        if elist[0] != 0:
            # Build message
            email = outlook.CreateItem(0) # email.olMailItem
            email.Subject = 'Inventory List'
            # Refine list of subscribers. Ommit 'Subscribe" from list
            for subscriber in arrListNoPrices:
                if subscriber == 'Subscribe':
                    continue
                else:
                    email.Recipients.Add(subscriber)
            # Attache File. Replace '/' for '\\' before sending path to outlook
            file = elist[0].replace('/', '\\\\')
            email.Attachments.Add(file)
            # Insert Signature Image on HTML body
            email.BodyFormat = 2
            email.HTMLBody = '<img width=624 height=267 src="' + MAIL_SIGNATURE_IMG + '" alt="Francisco_Signature" v:shapes="Picture_x0020_1">'
            # Send Message
            email.Send()
            print('*** List No Prices Sent Succesfully')
            counter += 1
            
        # Send List High Prices
        if elist[1] != 0:
            # Build message
            email = outlook.CreateItem(0) # email.olMailItem
            email.Subject = MAIL_SUBJECT + ' wp'
            # Refine list of subscribers. Ommit 'Subscribe" from list
            for subscriber in arrListHighPrices:
                if subscriber == 'Subscribe':
                    continue
                else:
                    email.Recipients.Add(subscriber)
            # Attache File. Replace '/' for '\\' before sending path to outlook
            file = elist[1].replace('/', '\\\\')
            email.Attachments.Add(file)
            # Insert Signature Image on HTML body
            email.BodyFormat = 2
            email.HTMLBody = '<img width=624 height=267 src="' + MAIL_SIGNATURE_IMG + '" alt="Francisco_Signature" v:shapes="Picture_x0020_1">'
            # Send Message
            email.Send()
            print('*** List High Prices Sent Succesfully')
            counter += 1
            
        # Send List Two Prices
        if elist[2] != 0:
            # Build message
            email = outlook.CreateItem(0) # email.olMailItem
            email.Subject = MAIL_SUBJECT + ' wp'
            # Refine list of subscribers. Ommit 'Subscribe" from list
            for subscriber in arrListTwoPrices:
                if subscriber == 'Subscribe':
                    continue
                else:
                    email.Recipients.Add(subscriber)
            # Attache File. Replace '/' for '\\' before sending path to outlook
            file = elist[2].replace('/', '\\\\')
            email.Attachments.Add(file)
            # Insert Signature Image on HTML body
            email.BodyFormat = 2
            email.HTMLBody = '<img width=624 height=267 src="' + MAIL_SIGNATURE_IMG + '" alt="Francisco_Signature" v:shapes="Picture_x0020_1">'
            # Send Message
            email.Send()
            print('*** List Two Prices Sent Succesfully')
            counter += 1
            
        # Send List Upc Prices
        if elist[3] != 0:
            # Build message
            email = outlook.CreateItem(0) # email.olMailItem
            email.Subject = MAIL_SUBJECT + ' wp'
            # Refine list of subscribers. Ommit 'Subscribe" from list
            for subscriber in arrListUpcPrices:
                if subscriber == 'Subscribe':
                    continue
                else:
                    email.Recipients.Add(subscriber)
            # Attache File. Replace '/' for '\\' before sending path to outlook
            file = elist[3].replace('/', '\\\\')
            email.Attachments.Add(file)
            # Insert Signature Image on HTML body
            email.BodyFormat = 2
            email.HTMLBody = '<img width=624 height=267 src="' + MAIL_SIGNATURE_IMG + '" alt="Francisco_Signature" v:shapes="Picture_x0020_1">'
            # Send Message
            email.Send()
            print('*** List UPC Prices Sent Succesfully')
            counter += 1
    
    except Exception as inst:
        print('Error on: sendListsViaOutlook:')
        print(inst)
        
    else:
        print('*** ' + str(counter) + ' Selected List(s) were succesfully sent !!!')
    
def closeForm():
    saveSettings()
    root.destroy()

def loadSettings():
    
    try:
        doc = xmlet.parse('settings.xml')
        data = doc.getroot()
        varIsQbgo.set(int(data.find('qbgo').text))
        varFileLocation.set(data.find('file').text)
        varExportFolder.set(data.find('folder').text)
        # Get Map Data
        mapping = data.find('map')
        varPartNumber.set(mapping.find('part').text)
        varDescription.set(mapping.find('description').text)
        varCost.set(mapping.find('cost').text)
        varPrice.set(mapping.find('price').text)
        varQtyHand.set(mapping.find('qtyOnHand').text)
        varQtySO.set(mapping.find('qtyOnSO').text)
        varQtyPO.set(mapping.find('qtyOnPO').text)
        varMpn.set(mapping.find('barcode').text)
        # Get Export Options
        exports = data.find('export')
        varChkNoPrices.set(int(exports.find('noPrices').text))
        varChkHighPrices.set(int(exports.find('highPrices').text))
        varChkTwoPrices.set(int(exports.find('twoPrices').text))
        varChkUpcPrices.set(int(exports.find('upcPrices').text))
        varChkWebPrices.set(int(exports.find('webPrices').text))
        varExportImages.set(int(exports.find('images').text))
        # Get Mail Lists
        listNoPrices = data.find('listNoPrices')
        for addr in listNoPrices.findall('address'):
            arrListNoPrices.append(addr.text)
        varListNoPrices.set(arrListNoPrices[0])
    
        listHighPrices = data.find('listHighPrices')
        for addr in listHighPrices.findall('address'):
            arrListHighPrices.append(addr.text)
        varListHighPrices.set(arrListHighPrices[0])
    
        listTwoPrices = data.find('listTwoPrices')
        for addr in listTwoPrices.findall('address'):
            arrListTwoPrices.append(addr.text)
        varListTwoPrices.set(arrListTwoPrices[0])
    
        listUpcPrices = data.find('listUpcPrices')
        for addr in listUpcPrices.findall('address'):
            arrListUpcPrices.append(addr.text)
        varListUpcPrices.set(arrListUpcPrices[0])
    
        listWebPrices = data.find('listWebPrices')
        for addr in listWebPrices.findall('address'):
            arrListWebPrices.append(addr.text)
        varListWebPrices.set(arrListWebPrices[0])
    
    except Exception as inst:
        print('Error on: loadSettings:')
        print(inst)
        

def saveSettings():
    
    try:
        data = xmlet.Element('data')
        qbgo = xmlet.SubElement(data, 'qbgo')
        qbgo.text = str(varIsQbgo.get())
        file = xmlet.SubElement(data, 'file')
        file.text = varFileLocation.get()
        folder = xmlet.SubElement(data, 'folder')
        folder.text = varExportFolder.get()
        # Set Map Data
        mapping = xmlet.SubElement(data, 'map')
        part = xmlet.SubElement(mapping, 'part')
        part.text = varPartNumber.get()
        description = xmlet.SubElement(mapping, 'description')
        description.text = varDescription.get()
        cost = xmlet.SubElement(mapping, 'cost')
        cost.text = varCost.get()
        price = xmlet.SubElement(mapping, 'price')
        price.text = varPrice.get()
        qtyOnHand = xmlet.SubElement(mapping, 'qtyOnHand')
        qtyOnHand.text = varQtyHand.get()
        qtyOnSO = xmlet.SubElement(mapping, 'qtyOnSO')
        qtyOnSO.text = varQtySO.get()
        qtyOnPO = xmlet.SubElement(mapping, 'qtyOnPO')
        qtyOnPO.text = varQtyPO.get()
        barcode = xmlet.SubElement(mapping, 'barcode')
        barcode.text = varMpn.get()
        # Set Export Options
        export = xmlet.SubElement(data, 'export')
        noPrices = xmlet.SubElement(export, 'noPrices')
        noPrices.text = str(varChkNoPrices.get())
        highPrices = xmlet.SubElement(export, 'highPrices')
        highPrices.text = str(varChkHighPrices.get())
        twoPrices = xmlet.SubElement(export, 'twoPrices')
        twoPrices.text = str(varChkTwoPrices.get())
        upcPrices = xmlet.SubElement(export, 'upcPrices')
        upcPrices.text = str(varChkUpcPrices.get())
        webPrices = xmlet.SubElement(export, 'webPrices')
        webPrices.text = str(varChkWebPrices.get())
        images = xmlet.SubElement(export, 'images')
        images.text = str(varExportImages.get())
        # Set Email Lists
        listNoPrices = xmlet.SubElement(data, 'listNoPrices')
        for address in arrListNoPrices:
            add = xmlet.SubElement(listNoPrices, 'address')
            add.text = address
    
        listHighPrices = xmlet.SubElement(data, 'listHighPrices')
        for address in arrListHighPrices:
            add = xmlet.SubElement(listHighPrices, 'address')
            add.text = address
    
        listTwoPrices = xmlet.SubElement(data, 'listTwoPrices')
        for address in arrListTwoPrices:
            add = xmlet.SubElement(listTwoPrices, 'address')
            add.text = address
    
        listUpcPrices = xmlet.SubElement(data, 'listUpcPrices')
        for address in arrListUpcPrices:
            add = xmlet.SubElement(listUpcPrices, 'address')
            add.text = address
    
        listWebPrices = xmlet.SubElement(data, 'listWebPrices')
        for address in arrListWebPrices:
            add = xmlet.SubElement(listWebPrices, 'address')
            add.text = address
        # Print to file
        tree = xmlet.ElementTree(data)
        tree.write('settings.xml')
        
    except Exception as inst:
        print('Error on: saveSettings:')
        print(inst)
        
    else:
        print('.... Settings Succesfully Saved.')

# Main Form
root = tkinter.Tk()
root.title('ILP - Inventory List Processor')
root.resizable(width=False, height=False)
root.iconbitmap(default='ilp_icon.ico')
frmMain = tkinter.Frame(root).grid(column=0, row=0)

# Form Variables
varIsQbgo = tkinter.IntVar()
varFileLocation = tkinter.StringVar()
varExportFolder = tkinter.StringVar()
varColumnLetters = ['NA', 'Column A', 'Column B', 'Column C', 'Column D', 'Column E', 'Column F', 'Column G', 'Column H']
varPartNumber = tkinter.StringVar()
varDescription = tkinter.StringVar()
varCost = tkinter.StringVar()
varPrice = tkinter.StringVar()
varQtyHand = tkinter.StringVar()
varQtySO = tkinter.StringVar()
varQtyPO = tkinter.StringVar()
varMpn = tkinter.StringVar()
varChkNoPrices = tkinter.IntVar()
varChkHighPrices = tkinter.IntVar()
varChkTwoPrices = tkinter.IntVar()
varChkUpcPrices = tkinter.IntVar()
varChkWebPrices = tkinter.IntVar()
varListNoPrices = tkinter.StringVar()
varListHighPrices = tkinter.StringVar()
varListTwoPrices = tkinter.StringVar()
varListUpcPrices = tkinter.StringVar()
varListWebPrices = tkinter.StringVar()
varExportImages = tkinter.IntVar()

arrListNoPrices = []
arrListHighPrices = []
arrListTwoPrices = []
arrListUpcPrices = []
arrListWebPrices = []

# Populate all variables before creating widgets
loadSettings()

# Form Widgets: 6 columns
# Row #0: Is QBgo File
ttk.Checkbutton(frmMain, text='QB Direct', variable=varIsQbgo, onvalue=1, offvalue=0).grid(column=5, row=0, padx=5, pady=5, sticky='W')
# Row #1-2: Files and Folders
ttk.Label(frmMain, text='File:', width=6).grid(column=0, row=1, padx=5, pady=5, sticky='W')
ttk.Entry(frmMain, textvariable=varFileLocation, width=60).grid(column=1, row=1, columnspan=4, padx=5, pady=5, sticky='WE')
ttk.Button(frmMain, text='Browse', command=browseFile, width=10).grid(column=5, row=1, padx=5, pady=5, sticky='E')
ttk.Label(frmMain, text='Folder:', width=6).grid(column=0, row=2, padx=5, pady=5, sticky='W')
ttk.Entry(frmMain, textvariable=varExportFolder, width=60).grid(column=1, row=2, columnspan=4, padx=5, pady=5, sticky='WE')
ttk.Button(frmMain, text='Browse', command=browseFolder, width=10).grid(column=5, row=2, padx=5, pady=5, sticky='E')
# Row #3: Separator
ttk.Label(frmMain, text='Map ', width=5).grid(column=0, row=3, padx=5, pady=5, sticky='W')
ttk.Separator(frmMain, orient=tkinter.constants.HORIZONTAL).grid(column=1, row=3, columnspan=5, padx=5, pady=5, sticky='WE')
# Row #4-7: Map
ttk.Label(frmMain, text='Part #:', width=10).grid(column=1, row=4, padx=5, pady=5, sticky='WE')
ttk.Combobox(frmMain, values=varColumnLetters, textvariable=varPartNumber, width=10).grid(column=2, row=4, padx=5, pady=5, sticky='W')
ttk.Label(frmMain, text='Qty on Hand:', width=12).grid(column=3, row=4, padx=5, pady=5, sticky='E')
ttk.Combobox(frmMain, values=varColumnLetters, textvariable=varQtyHand, width=10).grid(column=4, row=4, padx=5, pady=5, sticky='WE')
ttk.Label(frmMain, text='Description:', width=10).grid(column=1, row=5, padx=5, pady=5, sticky='WE')
ttk.Combobox(frmMain, values=varColumnLetters, textvariable=varDescription, width=10).grid(column=2, row=5, padx=5, pady=5, sticky='W')
ttk.Label(frmMain, text='Qty on S.O:', width=12).grid(column=3, row=5, padx=5, pady=5, sticky='E')
ttk.Combobox(frmMain, values=varColumnLetters, textvariable=varQtySO, width=10).grid(column=4, row=5, padx=5, pady=5, sticky='WE')
ttk.Label(frmMain, text='Cost:', width=10).grid(column=1, row=6, padx=5, pady=5, sticky='WE')
ttk.Combobox(frmMain, values=varColumnLetters, textvariable=varCost, width=10).grid(column=2, row=6, padx=5, pady=5, sticky='W')
ttk.Label(frmMain, text='Qty on P.O:', width=12).grid(column=3, row=6, padx=5, pady=5, sticky='E')
ttk.Combobox(frmMain, values=varColumnLetters, textvariable=varQtyPO, width=10).grid(column=4, row=6, padx=5, pady=5, sticky='WE')
ttk.Label(frmMain, text='Price:', width=10).grid(column=1, row=7, padx=5, pady=5, sticky='WE')
ttk.Combobox(frmMain, values=varColumnLetters, textvariable=varPrice, width=10).grid(column=2, row=7, padx=5, pady=5, sticky='W')
ttk.Label(frmMain, text='MPN:', width=12).grid(column=3, row=7, padx=5, pady=5, sticky='E')
ttk.Combobox(frmMain, values=varColumnLetters, textvariable=varMpn, width=10).grid(column=4, row=7, padx=5, pady=5, sticky='WE')
# Row #8: Separator
ttk.Label(frmMain, text='Exports', width=6).grid(column=0, row=8, padx=5, pady=5, sticky='W')
ttk.Separator(frmMain, orient=tkinter.constants.HORIZONTAL).grid(column=1, row=8, columnspan=5, padx=5, pady=5, sticky='WE')
# Row #9-13: Exports
ttk.Checkbutton(frmMain, text='List No Prices', variable=varChkNoPrices, onvalue=1, offvalue=0).grid(column=1, row=9, padx=5, pady=5, sticky='W')
cboNoPrices = ttk.Combobox(frmMain, values=arrListNoPrices, textvariable=varListNoPrices, width=20)
cboNoPrices.grid(column=2, row=9, padx=5, pady=5, sticky='WE')
ttk.Button(frmMain, text='Remove', command=lambda: modifyEmailList('lnp', 'delete'), width=3).grid(column=3, row=9, padx=5, pady=5, sticky='WE')
ttk.Button(frmMain, text='Add', command=lambda: modifyEmailList('lnp', 'add'), width=3).grid(column=4, row=9, padx=5, pady=5, sticky='WE')

ttk.Checkbutton(frmMain, text='List High Prices', variable=varChkHighPrices, onvalue=1, offvalue=0).grid(column=1, row=10, padx=5, pady=5, sticky='W')
cboHighPrices = ttk.Combobox(frmMain, values=arrListHighPrices, textvariable=varListHighPrices, width=20)
cboHighPrices.grid(column=2, row=10, padx=5, pady=5, sticky='WE')
ttk.Button(frmMain, text='Remove', command=lambda: modifyEmailList('lhp', 'delete'), width=3).grid(column=3, row=10, padx=5, pady=5, sticky='WE')
ttk.Button(frmMain, text='Add', command=lambda: modifyEmailList('lhp', 'add'), width=3).grid(column=4, row=10, padx=5, pady=5, sticky='WE')

ttk.Checkbutton(frmMain, text='List Two Prices', variable=varChkTwoPrices, onvalue=1, offvalue=0).grid(column=1, row=11, padx=5, pady=5, sticky='W')
cboTwoPrices = ttk.Combobox(frmMain, values=arrListTwoPrices, textvariable=varListTwoPrices, width=20)
cboTwoPrices.grid(column=2, row=11, padx=5, pady=5, sticky='WE')
ttk.Button(frmMain, text='Remove', command=lambda: modifyEmailList('l2p', 'delete'), width=3).grid(column=3, row=11, padx=5, pady=5, sticky='WE')
ttk.Button(frmMain, text='Add', command=lambda: modifyEmailList('l2p', 'add'), width=3).grid(column=4, row=11, padx=5, pady=5, sticky='WE')

ttk.Checkbutton(frmMain, text='List Price & UPC', variable=varChkUpcPrices, onvalue=1, offvalue=0).grid(column=1, row=12, padx=5, pady=5, sticky='W')
cboUpcPrices = ttk.Combobox(frmMain, values=arrListUpcPrices, textvariable=varListUpcPrices, width=20)
cboUpcPrices.grid(column=2, row=12, padx=5, pady=5, sticky='WE')
ttk.Button(frmMain, text='Remove', command=lambda: modifyEmailList('lup', 'delete'), width=3).grid(column=3, row=12, padx=5, pady=5, sticky='WE')
ttk.Button(frmMain, text='Add', command=lambda: modifyEmailList('lup', 'add'), width=3).grid(column=4, row=12, padx=5, pady=5, sticky='WE')

ttk.Checkbutton(frmMain, text='List Website', variable=varChkWebPrices, onvalue=1, offvalue=0).grid(column=1, row=13, padx=5, pady=5, sticky='W')
cboWebPrices = ttk.Combobox(frmMain, values=arrListWebPrices, textvariable=varListWebPrices, width=20)
cboWebPrices.grid(column=2, row=13, padx=5, pady=5, sticky='WE')
ttk.Button(frmMain, text='Remove', command=lambda: modifyEmailList('lwp', 'delete'), width=3).grid(column=3, row=13, padx=5, pady=5, sticky='WE')
ttk.Button(frmMain, text='Add', command=lambda: modifyEmailList('lwp', 'add'), width=3).grid(column=4, row=13, padx=5, pady=5, sticky='WE')
# Row #14: Separator
ttk.Separator(frmMain, orient=tkinter.constants.HORIZONTAL).grid(column=0, row=14, columnspan=6, padx=5, pady=5, sticky='WE')
# Row #15: Buttons
ttk.Checkbutton(frmMain, text='Export Missing Images', variable=varExportImages, onvalue=1, offvalue=0).grid(column=0, row=15, columnspan=2, padx=5, pady=5, sticky='W')
ttk.Button(frmMain, text='Ex & Send', command=sendListsViaOutlook, width=10).grid(column=3, row=15, padx=5, pady=5, sticky='WE')
ttk.Button(frmMain, text='Export', command=exportLists, width=10).grid(column=4, row=15, padx=5, pady=5, sticky='WE')
ttk.Button(frmMain, text='Close', command=closeForm, width=10).grid(column=5, row=15, padx=5, pady=5, sticky='WE')
# Show Form
root.mainloop()
