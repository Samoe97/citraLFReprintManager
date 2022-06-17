# CITRA COMMUNICATIONS
# LARGE FORMAT REPRINT MANAGER
# PYTHON v3.9.9
# Written by Samoe (www.samoe.me/code)

## IMPORTS ##

from tkinter import *
from tkinter.font import Font
from tkinter.ttk import Treeview, Style, Entry
from samoeModules import Hoverbutton, StickerTool2
import os, shutil
import datetime

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

  global requestRoot
  requestRoot = Toplevel()
  requestRoot.title('New Reprint Request')

  requestRootFrame = Frame(requestRoot)
  requestRootFrame.pack(padx = 16, pady = 16, ipadx = 4)

  requestTitle = Label(requestRootFrame, text = 'New Reprint Request', justify = 'left')
  requestTitle.grid(column = 0, row = 0, sticky = 'W')

  ###

  requestNameLabel = Label(requestRootFrame, text = 'User Name: ', justify = 'left')
  requestNameLabel.grid(column = 0, row = 1, pady = 4, sticky = 'W')

  global requestNameEntry
  requestNameEntry = Entry(requestRootFrame, width = 48)
  requestNameEntry.grid(column = 0, row = 2, columnspan = 2, pady = 4, ipady = 4, ipadx = 4, sticky = 'NEWS')
  
  ###

  requestOrderLabel = Label(requestRootFrame, text = 'Order Number: ', justify = 'left')
  requestOrderLabel.grid(column = 0, row = 3, pady = 4, sticky = 'W')

  global requestOrderEntry
  requestOrderEntry = Entry(requestRootFrame, width = 24)
  requestOrderEntry.grid(column = 0, row = 4, pady = 4, ipadx = 4, ipady = 4, sticky = 'NEWS')

  requestSkuLabel = Label(requestRootFrame, text = 'Sku: ', justify = 'left')
  requestSkuLabel.grid(column = 1, row = 3, pady = 4, sticky = 'W')

  global requestSkuEntry
  requestSkuEntry = Entry(requestRootFrame, width = 24)
  requestSkuEntry.grid(column = 1, row = 4, pady = 4, ipadx = 4, ipady = 4, sticky = 'NEWS')

  ###

  requestQtyNeededLabel = Label(requestRootFrame, text = 'Quantity Needed: ', justify = 'left')
  requestQtyNeededLabel.grid(column = 0, row = 5, pady = 4, sticky = 'W')

  global requestQtyNeededEntry
  requestQtyNeededEntry = Entry(requestRootFrame, width = 24)
  requestQtyNeededEntry.grid(column = 0, row = 6, pady = 4, ipadx = 4, ipady = 4, sticky = 'NEWS')

  requestQtyTotalLabel = Label(requestRootFrame, text = 'Quantity Total: ', justify = 'left')
  requestQtyTotalLabel.grid(column = 1, row = 5, pady = 4, sticky = 'W')

  global requestQtyTotalEntry
  requestQtyTotalEntry = Entry(requestRootFrame, width = 24)
  requestQtyTotalEntry.grid(column = 1, row = 6, pady = 4, ipadx = 4, ipady = 4, sticky = 'NEWS')

  ###

  requestSizeLabel = Label(requestRootFrame, text = 'Size: ', justify = 'left')
  requestSizeLabel.grid(column = 0, row = 7, pady = 4, sticky = 'W')

  global requestSizeEntry
  requestSizeEntry = Entry(requestRootFrame, width = 48)
  requestSizeEntry.grid(column = 0, row = 8, columnspan = 2, pady = 4, ipady = 4, ipadx = 4, sticky = 'NEWS')
  
  ###

  requestReasonLabel = Label(requestRootFrame, text = 'Reason: ', justify = 'left')
  requestReasonLabel.grid(column = 0, row = 9, pady = 4, sticky = 'W')

  global requestReasonEntry
  requestReasonEntry = Entry(requestRootFrame, width = 48)
  requestReasonEntry.grid(column = 0, row = 10, columnspan = 2, pady = 4, ipady = 4, ipadx = 4, sticky = 'NEWS')
  
  ###

  requestSubmitButton = Hoverbutton.HoverButton(requestRootFrame, root, text = 'Submit', command = submitRequest, colors = ['#426CB4', '#3762AE', '#284880'], fg = 'white')
  requestSubmitButton.grid(column = 0, row = 11, columnspan = 2, ipadx = 16, ipady = 16, sticky = 'NEWS', pady = 16)

  ###

  requestRoot.mainloop()

def submitRequest() :

  print('Submitting request')

  timestamp = datetime.datetime.now()
  timestamp = datetime.datetime.timestamp(timestamp)

  data = {
    'reason': requestReasonEntry.get(), 
    'itemId': requestSkuEntry.get(), 
    'updated': [''], 
    'status': 'requested', 
    'created': timestamp,
    'qtytotal': requestQtyTotalEntry.get(), 
    'qtyNeeded': requestQtyNeededEntry.get(), 
    'name': requestNameEntry.get(), 
    'size': requestSizeEntry.get(), 
    'orderId': requestOrderEntry.get(),
  }

  db.collection(u'reprints').document().set(data)
  print('Success')
  requestRoot.destroy()
  refresh()

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

    if str(docObject['created']).__contains__('.') :
      timeStamp = int(str(docObject['created'])[:-7])
    else : 
      timeStamp = int(str(docObject['created'])[:-3])
    timeStamp = datetime.datetime.fromtimestamp(timeStamp)
    timeStamp = timeStamp.strftime("%m/%d/%Y, %H:%M:%S")

    values = [docObject['orderId'], docObject['itemId'], docObject['size'], tempQty, docObject['name'], docObject['reason'], timeStamp, docObject['status']]

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
  root.config(bg = '#FFFFFF')
  rootFrame = Frame(root, bg = '#FFFFFF')
  rootFrame.grid(padx = 16, pady = 16, sticky = 'NEWS')

  fontMain = Font(family = 'Segoe ui', size = 12)

  root.grid_columnconfigure(0, weight = 1)
  root.grid_rowconfigure(0, weight = 1)

  rootFrame.grid_columnconfigure(0, weight = 1)
  rootFrame.grid_rowconfigure(0, weight = 10)
  rootFrame.grid_rowconfigure(1, weight = 1)
  rootFrame.grid_rowconfigure(2, weight = 10000)

  headerFrame = Frame(rootFrame, bg = '#FFFFFF')
  headerFrame.grid(row = 0, column = 0, sticky = 'NEWS')

  headerFrame.grid_columnconfigure(0, weight = 1000)
  headerFrame.grid_columnconfigure(1, weight = 1000)
  headerFrame.grid_columnconfigure(2, weight = 1)
  headerFrame.grid_rowconfigure(0, weight = 1)

  requestButton = Hoverbutton.HoverButton(headerFrame, root, text = 'New Reprint', command = newRequest, font = fontMain, colors = ['#158754', '#127448', '#156D44'], fg = 'white')
  requestButton.grid(column = 0, row = 0, ipadx = 16, ipady = 16, sticky = 'NEWS', padx = 16, pady = 16)

  pullButton = Hoverbutton.HoverButton(headerFrame, root, text = 'Pull Requested Reprints', command = pullFiles, font = fontMain)
  pullButton.grid(column = 1, row = 0, ipadx = 16, ipady = 16, sticky = 'NEWS', padx = 16, pady = 16)

  refreshButton = Hoverbutton.HoverButton(headerFrame, root, text = 'Refresh', command = refresh, font = fontMain)
  refreshButton.grid(column = 2, row = 0, ipadx = 16, ipady = 16, sticky = 'NEWS', padx = 16, pady = 16)

  ###

  seperatorFrame = Frame(rootFrame, bg = '#E6E6E6', height = 2)
  seperatorFrame.grid(column = 0, row = 1, sticky = 'NEWS', pady = 8)

  ###

  listFrame = Frame(rootFrame, bg = '#FFFFFF')
  listFrame.grid(row = 2, column = 0, sticky = 'NEWS')

  listFrame.grid_columnconfigure(0, weight = 1000)
  listFrame.grid_columnconfigure(1, weight = 1)
  listFrame.grid_rowconfigure(0, weight = 1)

  style = Style()
  style.configure("treeview", rowheight = 48, font = ('Segoe ui', 12))
  style.layout("treeview", [('mtreeviewStyle.treearea', {'sticky': 'NEWS'})])
  style.map('treeview', background = [('selected', '#F2F2F3')])

  style.configure("treeviewStyle",
                  background = "white",
                  foreground = "#111111",
                  rowheight = 25,
                  fieldbackground = "white")

  global reprintList

  reprintList = Treeview(listFrame, columns = ['Order', 'Sku', 'Size', 'Quantity', 'Submitter Name', 'Reason', 'Time', 'Status'], selectmode = 'extended', style = "treeview")
  reprintList.grid(column = 0, row = 0, sticky = 'NEWS')

  reprintList.heading('#0', text='', anchor = 'w')
  reprintList.heading('Order', text='Order', anchor = 'w')
  reprintList.heading('Sku', text ='Sku', anchor = 'w')
  reprintList.heading('Size', text = 'Size', anchor = 'w')
  reprintList.heading('Quantity', text = 'Quantity', anchor = 'w')
  reprintList.heading('Submitter Name', text='Submitter Name', anchor = 'w')
  reprintList.heading('Reason', text='Reason', anchor = 'w')
  reprintList.heading('Time', text = 'Time', anchor = 'w')
  reprintList.heading('Status', text='Status', anchor = 'w')

  reprintList.column('#0', width = 16, minwidth = 16, anchor = 'w')
  reprintList.column('Order', width = 200, minwidth = 96, anchor = 'w')
  reprintList.column('Sku', width = 200, minwidth = 96, anchor = 'w')
  reprintList.column('Size', width = 200, minwidth = 96, anchor = 'w')
  reprintList.column('Quantity', width = 200, minwidth = 96, anchor = 'w')
  reprintList.column('Submitter Name', width = 300, minwidth = 96, anchor = 'w')
  reprintList.column('Reason', width = 200, minwidth = 96, anchor = 'w')
  reprintList.column('Time', width = 200, minwidth = 96, anchor = 'w')
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