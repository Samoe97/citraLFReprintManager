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

def newRequest() :

  print('Opening new reprint request window.')

  requestRoot = Toplevel()
  requestRoot.title('New Reprint Request')

  requestRootFrame = Frame(requestRoot)
  requestRootFrame.pack(padx = 16, pady = 16, ipadx = 4)

  requestTitle = Label(requestRootFrame, text = 'New Reprint Request', justify = 'left')
  requestTitle.grid(column = 0, row = 0, sticky = 'W')

  ###

  requestNameLabel = Label(requestRootFrame, text = 'User Name: ', justify = 'left')
  requestNameLabel.grid(column = 0, row = 1, pady = 4, sticky = 'W')

  requestNameEntry = Entry(requestRootFrame, width = 48)
  requestNameEntry.grid(column = 0, row = 2, columnspan = 2, pady = 4, ipady = 4, ipadx = 4, sticky = 'NEWS')
  
  ###

  requestOrderLabel = Label(requestRootFrame, text = 'Order Number: ', justify = 'left')
  requestOrderLabel.grid(column = 0, row = 3, pady = 4, sticky = 'W')

  requestOrderEntry = Entry(requestRootFrame, width = 24)
  requestOrderEntry.grid(column = 0, row = 4, pady = 4, ipadx = 4, ipady = 4, sticky = 'NEWS')

  requestSkuLabel = Label(requestRootFrame, text = 'Sku: ', justify = 'left')
  requestSkuLabel.grid(column = 1, row = 3, pady = 4, sticky = 'W')

  requestSkuEntry = Entry(requestRootFrame, width = 24)
  requestSkuEntry.grid(column = 1, row = 4, pady = 4, ipadx = 4, ipady = 4, sticky = 'NEWS')

  ###

  requestRoot.mainloop()

def refresh() :

  print('Refreshing')

  reprintsRef = db.collection('reprints')
  docs = reprintsRef.stream()

  for item in reprintList.get_children():
        reprintList.delete(item)

  for doc in docs:
    docObject = doc.to_dict()
    docObject.update({'id' : doc.id})

    tempQty = f'{docObject["qtyNeeded"]} / {docObject["qtytotal"]}'

    values = [docObject['orderId'], docObject['itemId'], docObject['size'], tempQty, docObject['name'], docObject['reason'], docObject['status']]

    reprintList.insert('', 'end',
        values = values, tags = (docObject['status']))

    reprintList.tag_configure('deleted', background=colors['deleted'])
    reprintList.tag_configure('complete', background=colors['complete'])
    reprintList.tag_configure('requested', background=colors['requested'])
    reprintList.tag_configure('pulled', background=colors['pulled'])
    reprintList.tag_configure('printed', background=colors['printed'])

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

  refresh()

##############################

def startApp() :

  print('Start')

  global root
  root = Tk()
  root.title('Citra Communications - Large Format Reprint Manager')
  rootFrame = Frame(root)
  rootFrame.grid(padx = 16, pady = 16, sticky = 'NEWS')

  root.grid_columnconfigure(0, weight = 1)
  root.grid_rowconfigure(0, weight = 1)

  rootFrame.grid_columnconfigure(0, weight = 1)
  rootFrame.grid_rowconfigure(0, weight = 1)
  rootFrame.grid_rowconfigure(1, weight = 1000)

  headerFrame = Frame(rootFrame)
  headerFrame.grid(row = 0, column = 0, sticky = 'NEWS')

  headerFrame.grid_columnconfigure(0, weight = 1)
  headerFrame.grid_columnconfigure(1, weight = 1)
  headerFrame.grid_columnconfigure(2, weight = 1)
  headerFrame.grid_rowconfigure(0, weight = 1)

  requestButton = Hoverbutton.HoverButton(headerFrame, root, text = 'New Reprint', command = newRequest)
  requestButton.grid(column = 0, row = 0, ipadx = 16, ipady = 16, sticky = 'NEWS', padx = 16, pady = 16)

  pullButton = Hoverbutton.HoverButton(headerFrame, root, text = 'Pull Requested Reprints', command = pullFiles)
  pullButton.grid(column = 1, row = 0, ipadx = 16, ipady = 16, sticky = 'NEWS', padx = 16, pady = 16)

  refreshButton = Hoverbutton.HoverButton(headerFrame, root, text = 'Refresh', command = refresh)
  refreshButton.grid(column = 2, row = 0, ipadx = 16, ipady = 16, sticky = 'NEWS', padx = 16, pady = 16)

  listFrame = Frame(rootFrame)
  listFrame.grid(row = 1, column = 0, sticky = 'NEWS')

  listFrame.grid_columnconfigure(0, weight = 1000)
  listFrame.grid_columnconfigure(1, weight = 1)
  listFrame.grid_rowconfigure(0, weight = 1)

  style = Style()
  style.configure("treeview", rowheight = 48, font=('Segoe ui', 12))
  style.layout("treeview", [('mtreeviewStyle.treearea', {'sticky': 'NEWS'})])
  style.map('treeview', background=[('selected', '#BFBFBF')])

  style.configure("treeviewStyle",
                  background="#E1E1E1",
                  foreground="#000000",
                  rowheight=25,
                  fieldbackground="#E1E1E1")

  global reprintList

  reprintList = Treeview(listFrame, columns = ['Order', 'Sku', 'Size', 'Quantity', 'Submitter Name', 'Reason', 'Status'], selectmode = 'extended', style="treeview")
  reprintList.grid(column = 0, row = 0, sticky = 'NEWS')

  reprintList.heading('#0', text='', anchor = 'w')
  reprintList.heading('Order', text='Order', anchor = 'w')
  reprintList.heading('Sku', text ='Sku', anchor = 'w')
  reprintList.heading('Size', text = 'Size', anchor = 'w')
  reprintList.heading('Quantity', text = 'Quantity', anchor = 'w')
  reprintList.heading('Submitter Name', text='Submitter Name', anchor = 'w')
  reprintList.heading('Reason', text='Reason', anchor = 'w')
  reprintList.heading('Status', text='Status', anchor = 'w')

  reprintList.column('#0', width = 16, minwidth = 16, anchor = 'w')
  reprintList.column('Order', width = 200, minwidth = 96, anchor = 'w')
  reprintList.column('Sku', width = 200, minwidth = 96, anchor = 'w')
  reprintList.column('Size', width = 200, minwidth = 96, anchor = 'w')
  reprintList.column('Quantity', width = 200, minwidth = 96, anchor = 'w')
  reprintList.column('Submitter Name', width = 300, minwidth = 96, anchor = 'w')
  reprintList.column('Reason', width = 200, minwidth = 96, anchor = 'w')
  reprintList.column('Status', width = 200, minwidth = 96, anchor = 'w')
  
  scrollbar = Scrollbar(listFrame)
  scrollbar.grid(column = 1, row = 0, sticky = 'NEWS')

  reprintList.configure(yscrollcommand = scrollbar.set)
  scrollbar.config(command = reprintList.yview)
  # reprintList.bind('<Button-1>', selectItem)

  refresh()

  root.mainloop()

if __name__ == '__main__' :
  startApp()