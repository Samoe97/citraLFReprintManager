# CITRA COMMUNICATIONS
# LARGE FORMAT REPRINT MANAGER
# PYTHON v3.9.9
# Written by Samoe (www.samoe.me/code)

## IMPORTS ##

from tkinter import *
from tkinter.ttk import Treeview, Style
from samoeModules import Hoverbutton, StickerTool2
import os, shutil

## STYLE OPTIONS ##

colors = {
  "requested" : "#FFF4CE",
  "pulled" : "#D1EDF6",
  "printed" : "#D1E1F3",
  "complete" : "#D1E7DC",
  "deleted" : "#F7D7DA"
}

## FIRESTORE INITIALIZATION ##

import firebase_admin
from firebase_admin import credentials
from firebase_admin import firestore

credJson = os.path.dirname(__file__) + '/citralargeformatreprints-firebase-adminsdk-ea431-ad4ea283d6.json'

cred = credentials.Certificate(credJson)

firebase_admin.initialize_app(cred, {
  'projectId': 'citralargeformatreprints',
})

db = firestore.client()

##############################

def pullFiles() :

  print('Pulling')

  destFilePath = 'L:/Incoming Files/REPRINTS/'
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
    
    orderNumber = i['orderId']
    sku = i['itemId']
    size = i['size']
    qtyNeeded = i['qtyNeeded']
    qtyTotal = i['qtytotal']

    printFileName = f'{orderNumber}_{sku}_{size}_qty {qtyTotal}_PRINT.pdf'
    ticketFileName = f'{orderNumber}_{sku}_{size}_qty {qtyTotal}_TICKET.pdf'

    printFilePath = f'L:/Archive/TechStyles Archive/{orderNumber[0:2]}000 - {orderNumber[0:2]}999/{orderNumber}/{printFileName}'
    ticketFilePath = f'L:/Archive/TechStyles Archive/{orderNumber[0:2]}000 - {orderNumber[0:2]}999/{orderNumber}/{ticketFileName}'

    try :
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

    except :
      print('Error: Failed to copy file.')

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
  pullButton.grid(ipadx = 16, ipady = 16)

  listFrame = Frame(rootFrame)
  listFrame.grid(row = 1)

  style = Style()
  style.configure("treeviewStyle", rowheight = 32, font=('Roboto', 11))
  style.layout("treeviewStyle", [('mtreeviewStyle.treearea', {'sticky': 'NEWS'})])
  style.map('treeviewStyle', background=[('selected', '#BFBFBF')])

  style.configure("treeviewStyle",
                  background="#E1E1E1",
                  foreground="#000000",
                  rowheight=25,
                  fieldbackground="#E1E1E1")

  reprintList = Treeview(listFrame, columns = ['Order', 'Sku', 'Submitter Name', 'Reason', 'Status'], selectmode = 'extended', style="treeviewStyle")
  reprintList.grid(column = 0, row = 1, sticky = 'NEWS')

  reprintList.heading('#0', text='', anchor = 'w')
  reprintList.heading('Order', text='Order', anchor = 'w')
  reprintList.heading('Sku', text ='Sku', anchor = 'w')
  reprintList.heading('Submitter Name', text='Submitter Name', anchor = 'w')
  reprintList.heading('Reason', text='Reason', anchor = 'w')
  reprintList.heading('Status', text='Status', anchor = 'w')

  reprintList.column('#0', width = 128, minwidth = 96, anchor = 'w')
  reprintList.column('Order', width = 128, minwidth = 96, anchor = 'w')
  reprintList.column('Sku', width = 128, minwidth = 96, anchor = 'w')
  reprintList.column('Submitter Name', width = 128, minwidth = 96, anchor = 'w')
  reprintList.column('Reason', width = 96, minwidth = 96, anchor = 'w')
  reprintList.column('Status', width = 96, minwidth = 96, anchor = 'w')
  
  scrollbar = Scrollbar(listFrame)
  scrollbar.grid(column = 1, row = 1, sticky = 'NEWS')

  reprintList.configure(yscrollcommand = scrollbar.set)
  scrollbar.config(command = reprintList.yview)
  # reprintList.bind('<Button-1>', selectItem)

  reprintsRef = db.collection('reprints')
  docs = reprintsRef.stream()

  for doc in docs:
    docObject = doc.to_dict()
    docObject.update({'id' : doc.id})

    values = [docObject['orderId'], docObject['itemId'], docObject['name'], docObject['reason'], docObject['status']]

    reprintList.insert('', 'end',
        values = values, tags = (docObject['status']))

    reprintList.tag_configure('deleted', background=colors['deleted'])
    reprintList.tag_configure('complete', background=colors['complete'])
    reprintList.tag_configure('requested', background=colors['requested'])
    reprintList.tag_configure('pulled', background=colors['pulled'])
    reprintList.tag_configure('printed', background=colors['printed'])


  root.mainloop()

if __name__ == '__main__' :
  startApp()