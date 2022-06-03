# CITRA COMMUNICATIONS
# LARGE FORMAT REPRINT MANAGER
# PYTHON v3.9.9
# Written by Samoe (www.samoe.me/code)

## IMPORTS ##

from tkinter import *
from tkinter.ttk import Treeview
from samoeModules import Hoverbutton, StickerTool2
import os, shutil

# FIRESTORE INITIALIZATION ##

import firebase_admin
from firebase_admin import credentials
from firebase_admin import firestore

cred = credentials.Certificate('D:/Drive/Coding/citraLFReprintManager/samoeflasktest-c82acd1b6738.json')
firebase_admin.initialize_app(cred, {
  'projectId': 'samoeflasktest',
})

db = firestore.client()

##############################

def pullFiles() :

  print('Pulling')

  # destFilePath = 'L:/Incoming Files/REPRINTS/workfolder/'
  destFilePath = 'D:/Drive/Coding/citraLFReprintManager/destFolder/'
  filesToGet = []

  # GET LIST OF REPRINTS

  reprintsRef = db.collection('reprints')
  docs = reprintsRef.stream()

  for doc in docs:
    docObject = doc.to_dict()
    docObject.update({'id' : doc.id})

    if docObject['status'] == 'requested' :
      filesToGet.append(docObject)

  for i in filesToGet :
    
    orderNumber = i['order']
    sku = i['sku']
    size = i['size']
    qtyNeeded = i['qtyNeeded']
    qtyTotal = i['qtyTotal']

    printFileName = f'{orderNumber}_{sku}_{size}_qty {qtyTotal}_PRINT.pdf'
    ticketFileName = f'{orderNumber}_{sku}_{size}_qty {qtyTotal}_TICKET.pdf'

    # printFilePath = f'L:/Archive/TechStyles Archive/{orderNumber[0:2]}000 - {orderNumber[0:2]}999/{orderNumber}/{printFileName}'
    # ticketFilePath = f'L:/Archive/TechStyles Archive/{orderNumber[0:2]}000 - {orderNumber[0:2]}999/{orderNumber}/{ticketFileName}'

    printFilePath = f'D:/Drive/Coding/citraLFReprintManager/testFiles/{orderNumber[0:2]}000 - {orderNumber[0:2]}999/{orderNumber}/{printFileName}'
    ticketFilePath = f'D:/Drive/Coding/citraLFReprintManager/testFiles/{orderNumber[0:2]}000 - {orderNumber[0:2]}999/{orderNumber}/{ticketFileName}'

    # try :
    shutil.copy(printFilePath, destFilePath + printFileName)
    shutil.copy(ticketFilePath, destFilePath + ticketFileName)
  
    newPrintFile = f'{orderNumber}_{sku}_{size}_qty {qtyNeeded}_PRINT.pdf'
    newTicketFile = f'{orderNumber}_{sku}_{size}_qty {qtyNeeded}_TICKET.pdf'

    os.rename(destFilePath + printFileName, destFilePath + newPrintFile)
    os.rename(destFilePath + ticketFileName, destFilePath + newTicketFile)

    doc_ref = db.collection('reprints').document(i['id'])
    doc_ref.update({
      'status': 'pulled'
    })

    # except :
      # print('Error: Failed to copy file.')

  # GENERATE STICKER SHEETS
  params = {
          'Directory' : destFilePath,
          'ExtraStickers' : 4,
          'SpaceBetweenStickers' : '0.125',
          'MaxSheetHeight' : 32,
          'MaxSheetWidth' : 50,
          'DontIncludeTicket' : False,
          'TicketOnlyOnFirstSheet' : False,
          'Archive1UPs' : False,
          'addCutLineToTicket' : False
      }
  StickerTool2.executeStickerSheetGenerator(params)

##############################

def startApp() :

  print('Start')

  global root
  root = Tk()
  rootFrame = Frame(root)
  rootFrame.grid(padx = 16, pady = 16, sticky = 'NEWS')

  root.grid_columnconfigure(0, weight = 1)
  root.grid_rowconfigure(0, weight = 1)

  pullButton = Hoverbutton.HoverButton(rootFrame, root, text = 'Pull Requested Reprints', command = pullFiles)
  pullButton.pack(ipadx = 16, ipady = 16)

  root.mainloop()

if __name__ == '__main__' :
  startApp()