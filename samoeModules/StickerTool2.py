# STICKER TOOL V2 FROM CITRABOX
# SAMOE.ME/CODE

# MODIFIED FOR CITRA COMMUNICATIONS REPRINT MANAGER

import os
import math

from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.colors import CMYKColorSep

from pdfrw import PdfReader
from pdfrw.buildxobj import pagexobj
from pdfrw.toreportlab import makerl

global sheetIndex
sheetIndex = 1

def points(inches):
    points = float(inches) * 72
    return float(points)

def executeStickerSheetGenerator(params) :
    
    allPrintPDFs = []
    allInfoPDFs = []
    allBarcodes = []

    for file in os.listdir(params['Directory']):
        if file.endswith('PRINT.pdf') and not file.__contains__('TILE'):
            allPrintPDFs.append(file)

        elif file.endswith('TICKET.pdf') and not file.__contains__('TILE'):
            if params['DontIncludeTicket'] == False :
                allInfoPDFs.append(file)

    if params['addCutLineToTicket'] == True:
        for i in allInfoPDFs :
            itemSpecs = fileNameParser(i)

            filepath = params['Directory'] + '/' + i
            pythonInfoCut(filepath, itemSpecs)

    # Create Sticker Sheets
    for i in allPrintPDFs :
        index = allPrintPDFs.index(i)
        itemSpecs = fileNameParser(i)
        if params['DontIncludeTicket'] == True:
            infotech = ''
        else :
            infotech = allInfoPDFs[index]

        pythonSheetGenerator(i, infotech, itemSpecs, params)

    print(' ---- Sticker sheet generation complete! ---- \n')
    
def fileNameParser(filename):
    width = 0
    height = 0
    filename = str(filename).split('/')[0]
    if filename.startswith('TILE') :
        orderNumber = filename.split('_')[1]
        skuNumber = filename.split('_')[2]
        if skuNumber.__contains__('-') :
            skuNumber.replace('-', '')
    elif filename.startswith('ROUND') :
        orderNumber = filename.split('_')[1]
        skuNumber = filename.split('_')[2]
    else:
        orderNumber = filename[0:5]
        skuNumber = filename[6:12]
        skuNumber = skuNumber.split('_')[0]
        size = str(filename).split('x')
        width = size[0][-1]
        height = size[1].split('_')[0]
    quantityfile = str(filename).split('_qty ')[1]
    quantityfile = str(quantityfile).split('_')[0]

    if quantityfile.__contains__('.pdf') :
        quantityfile = quantityfile[:-4]

    itemSpecs = {
        'order' : orderNumber,
        'sku' : skuNumber,
        'width' : width,
        'height' : height,
        'quantity' : quantityfile
    }
    return itemSpecs

def pythonInfoCut(infotech, itemSpecs):

    ticketFile = PdfReader(infotech).pages[0]
    ticketFile = pagexobj(ticketFile)

    width = points(itemSpecs["width"][0]),
    height = points(itemSpecs["height"][0]),

    width = float(width[0])
    height = float(height[0])

    canvas = Canvas(infotech)
    canvas.setPageSize((width, height))

    canvas.doForm(makerl(canvas, ticketFile))

    width = width - points(0.2)
    height = height - points(0.2)
    positionX = points(0.1)
    positionY = points(0.1)

    perfCut = CMYKColorSep(1, 0, 1, 0, 'PerfCutContour', 1, 1)

    canvas.setStrokeColor(perfCut)

    canvas.rect(positionX, positionY, width, height, stroke=1, fill=0)

    canvas.save()

def pythonSheetGenerator(printFileName, ticketFileName, itemSpecs, params) :
    
    printFile = PdfReader(params['Directory'] + '/' + printFileName).pages[0]
    printFile = pagexobj(printFile)

    if ticketFileName != '' :
        ticketFile = PdfReader(params['Directory'] + '/' + ticketFileName).pages[0]
        ticketFile = pagexobj(ticketFile)
    else : ticketFile = printFile

    global sheetIndex
    global spaceBetweenStickers

    width = points(itemSpecs["width"])
    height = points(itemSpecs["height"])
    space = points(params['SpaceBetweenStickers'])

    printQuantity = int(itemSpecs["quantity"]) + int(params['ExtraStickers'])

    rows = math.ceil(points(params['MaxSheetHeight']) / (height + space))
    columns = math.floor(points(params['MaxSheetWidth']) / (width + space))

    canvasWidth = (width + space) * columns
    canvasHeight = (height + space) * rows

    qtyPerSheet = rows * columns

    if qtyPerSheet >= printQuantity :
        qtyPerSheet = printQuantity

    sheetsNeeded = math.ceil(printQuantity / qtyPerSheet)


    # MAIN SHEET CREATION LOOP
    for i in range(sheetsNeeded) :

        print(' - Creating sheet ' + str(sheetIndex) + ' of ' + str(sheetsNeeded) + ' for order ' + str(itemSpecs["order"]) + ' | SKU ' + str(itemSpecs["sku"]))
       
        # RECALCULATE SHEET HEIGHT IF THIS IS THE LAST SHEET
        if i + 1 == sheetsNeeded :
            rows = math.ceil(printQuantity / columns)
            canvasHeight = (height + space) * rows

        # INITIALIZE THE PDF WE'RE CREATING
        destination = params['Directory'] + '/' + printFileName[:-4] + '_Sheet' + str(sheetIndex) + '.pdf'
        canvas = Canvas(destination)    # THIS IS WHERE YOU CHOOSE WHERE THE FILE WILL SAVE TO
        canvas.setPageSize((canvasWidth, canvasHeight))

        # INITIAL STICKER POSITION (USUALLY THE INFOTECH SLOT)
        Xposition = 0
        Yposition = canvasHeight - (height)

        # LOOP THAT PUTS STICKERS ONTO SHEET
        for f in range(qtyPerSheet) :

            # OPEN CANVAS TO BE EDITED
            canvas.saveState()

            # IF WE REACH THE LAST COLUMN, MOVE TO THE NEXT ROW
            if f % columns == 0 and f != 0 :
                Xposition = 0
                Yposition = Yposition - (height + space)
            
            # SET POSITION OF NEW STICKER
            canvas.translate(Xposition, Yposition)
            Xposition = Xposition + (width + space)

            # PLACE TICKET OR PRINT FILE
            if params['TicketOnlyOnFirstSheet'] == True:     # THIS IS FOR THE OPTION OF ONLY HAVING THE TICKET ON THE FIRST SHEET
                if sheetIndex == 1:
                    canvas.doForm(makerl(canvas, ticketFile))
                    printQuantity = printQuantity + 1
                else:
                    canvas.doForm(makerl(canvas, printFile))

            else :
                if f == 0:  # IF THIS IS THE FIRST SLOT ON THE SHEET, PLACE A TICKET
                    if params['DontIncludeTicket'] == False:
                        canvas.doForm(makerl(canvas, ticketFile))
                        printQuantity = printQuantity + 1
                    else: 
                        canvas.doForm(makerl(canvas, printFile))    # IF THE USER CHOOSES TO NOT INCLUDE TICKETS, PLACE A PRINT FILE INSTEAD
                
                # PLACE A PRINT FILE
                else:
                    canvas.doForm(makerl(canvas, printFile))

            # CLOSE CANVAS, SIGNIFYING WE ARE DONE EDITING
            canvas.restoreState()

 
        printQuantity = printQuantity - qtyPerSheet     # PROCESS QUANTITY
        canvas.save()   # SAVE FILE (DESTINATION IS SET WHEN WE INITIALIZED CANVAS)
        sheetIndex = sheetIndex + 1     # ADD 1 TO THE SHEET COUNT, LOOP BACK TO BEGINNING IF NEEDED

    sheetIndex = 1

    os.remove(params['Directory'] + '/' + printFileName)
    os.remove(params['Directory'] + '/' + ticketFileName)

    # oneUpsPath = params['Directory'] + '/1UPs/'
    # if not os.path.isdir(oneUpsPath):
    #     os.mkdir(oneUpsPath)
    # newPrintFilePath = oneUpsPath + printFileName
    # newTicketFilePath = oneUpsPath + ticketFileName
    # shutil.move(params['Directory'] + '/' + printFileName, newPrintFilePath)
    # if params['DontIncludeTicket'] == False:
    #     shutil.move(params['Directory'] + '/' + ticketFileName, newTicketFilePath)
