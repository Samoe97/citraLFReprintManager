# CITRA COMMUNICATIONS
# LARGE FORMAT REPRINT MANAGER
# PYTHON v3.9.9
# Written by Samoe (www.samoe.me/code)

## IMPORTS ##

from tkinter import *
from tkinter.font import Font
from tkinter.ttk import Treeview, Style, Entry, Combobox
import os, shutil
import datetime

from samoeModules import Hoverbutton, StickerTool2

## EXTRA IMPORTS FOR COMPILING ##
import math
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.colors import CMYKColorSep
from pdfrw import PdfReader
from pdfrw.buildxobj import pagexobj
from pdfrw.toreportlab import makerl
import time
from tkinter import Label

filters = []
reprintListRef = []

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
  requestNameEntry.grid(column = 0, row = 2, columnspan = 3, pady = 4, ipady = 4, ipadx = 4, sticky = 'NEWS')
  requestNameEntry.focus()
  
  ###

  requestOrderLabel = Label(requestRootFrame, text = 'Order Number: ', justify = 'left')
  requestOrderLabel.grid(column = 0, row = 3, pady = 4, sticky = 'W')

  global requestOrderEntry
  requestOrderEntry = Entry(requestRootFrame, width = 24)
  requestOrderEntry.grid(column = 0, row = 4, pady = 4, ipadx = 4, ipady = 4, sticky = 'NEWS')

  requestSpacer1 = Frame(requestRootFrame, width = 1)
  requestSpacer1.grid(column = 1, row = 3, rowspan = 5, padx = 8)

  requestSkuLabel = Label(requestRootFrame, text = 'Sku: ', justify = 'left')
  requestSkuLabel.grid(column = 2, row = 3, pady = 4, sticky = 'W')

  global requestSkuEntry
  requestSkuEntry = Entry(requestRootFrame, width = 24)
  requestSkuEntry.grid(column = 2, row = 4, pady = 4, ipadx = 4, ipady = 4, sticky = 'NEWS')

  ###

  requestQtyNeededLabel = Label(requestRootFrame, text = 'Quantity Needed: ', justify = 'left')
  requestQtyNeededLabel.grid(column = 0, row = 5, pady = 4, sticky = 'W')

  global requestQtyNeededEntry
  requestQtyNeededEntry = Entry(requestRootFrame, width = 24)
  requestQtyNeededEntry.grid(column = 0, row = 6, pady = 4, ipadx = 4, ipady = 4, sticky = 'NEWS')

  requestQtyTotalLabel = Label(requestRootFrame, text = 'Quantity Total: ', justify = 'left')
  requestQtyTotalLabel.grid(column = 2, row = 5, pady = 4, sticky = 'W')

  global requestQtyTotalEntry
  requestQtyTotalEntry = Entry(requestRootFrame, width = 24)
  requestQtyTotalEntry.grid(column = 2, row = 6, pady = 4, ipadx = 4, ipady = 4, sticky = 'NEWS')

  ###

  requestSizeLabel = Label(requestRootFrame, text = 'Size: ', justify = 'left')
  requestSizeLabel.grid(column = 0, row = 7, pady = 4, sticky = 'W')

  global requestSizeEntry
  requestSizeEntry = Entry(requestRootFrame, width = 48)
  requestSizeEntry.grid(column = 0, row = 8, columnspan = 3, pady = 4, ipady = 4, ipadx = 4, sticky = 'NEWS')
  
  ###

  requestReasonLabel = Label(requestRootFrame, text = 'Reason: ', justify = 'left')
  requestReasonLabel.grid(column = 0, row = 9, pady = 4, sticky = 'W')

  global requestReasonEntry
  requestReasonEntry = Entry(requestRootFrame, width = 48)
  requestReasonEntry.grid(column = 0, row = 10, columnspan = 3, pady = 4, ipady = 4, ipadx = 4, sticky = 'NEWS')
  
  ###

  requestSubmitButton = Hoverbutton.HoverButton(requestRootFrame, root, text = 'Submit', command = submitRequest, colors = ['#426CB4', '#3762AE', '#284880'], fg = 'white')
  requestSubmitButton.grid(column = 0, row = 11, columnspan = 3, ipadx = 16, ipady = 16, sticky = 'NEWS', pady = 16)

  ###

  requestRoot.mainloop()

#############################################

def submitRequest() :

  print('Attempting to submit reprint...')

  timestamp = datetime.datetime.now()
  timestamp = datetime.datetime.timestamp(timestamp)

  if requestReasonEntry.get() == '' :
    print('Error: Reason required.')
    return
  
  if requestSkuEntry.get() == '' :
    print('Error: Sku required.')
    return

  if requestOrderEntry.get() == '' or len(requestOrderEntry.get()) <= 4 :
    print('Error: Order required.')
    return

  if requestNameEntry.get() == '' :
    print('Error: Name required.')
    return
  
  if requestSizeEntry.get() == '' :
    print('Error: Size required.')
    return

  if not len(requestSizeEntry.get()) == 3 :
    print('Error: Size must be typed in correctly, like 3x4 or 6x5.')
    return

  if not requestSizeEntry.get().__contains__('x') :
    print('Error: Size must contain an x between the two numbers, like 4x4 or 3x2.')
    return

  if requestQtyNeededEntry.get() == '' :
    print('Error: Quantity required.')
    return

  if requestQtyTotalEntry.get() == '' :
    print('Error: Total quantity required.')
    return

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
  print(f'Submitted new reprint: Order {requestOrderEntry.get()} | SKU {requestSkuEntry.get()}')
  requestRoot.destroy()
  refresh()

########################################

def changeStatusDialogue() :

  print('Opening window to change reprint status.')

  global statusRoot
  statusRoot = Toplevel()
  statusRoot.title('Change Reprint Status')

  statusRootFrame = Frame(statusRoot)
  statusRootFrame.pack(padx = 16, pady = 16, ipadx = 4)

  statusTitle = Label(statusRootFrame, text = 'Change Reprint Status', justify = 'center')
  statusTitle.grid(column = 0, row = 0, sticky = 'NEWS', pady = 8)

  ###

  global statusDropdown
  statusDropdown = Combobox(statusRootFrame, values = ['requested', 'pulled', 'deleted', 'printed', 'complete'], state = 'readonly', exportselection = True, justify = 'center')
  statusDropdown.grid(column = 0, row = 1, sticky = 'NEWS', pady = 8, ipady = 8, ipadx = 8)
  statusDropdown.current(0)

  ###

  statusSubmitButton = Hoverbutton.HoverButton(statusRootFrame, root, text = 'Submit', command = submitChangeStatus, colors = ['#426CB4', '#3762AE', '#284880'], fg = 'white')
  statusSubmitButton.grid(column = 0, row = 2, ipadx = 16, ipady = 16, sticky = 'NEWS')

  ###

  statusRoot.mainloop()

#############################################

def submitChangeStatus() :

  selectedItems = reprintList.selection()

  for i in selectedItems :

    reprint = reprintList.item(i)

    for a in reprintListRef :
      
      if str(reprint['values'][0]) == a['orderId'] :
        if a['itemId'].__contains__(str(reprint['values'][1])) :
        # if str(reprint['values'][1]) == a['itemId'] :
          if str(reprint['values'][2]) == a['size'] :
            doc_ref = db.collection('reprints').document(a['id'])
            doc_ref.update({
              'status': statusDropdown.get()
            })

            print('Success - Changed Order ' + a['orderId'] + ' | SKU ' + a['itemId'] + ' to ' + statusDropdown.get())

  refresh()
  statusRoot.destroy()

#############################################

def refresh() :

  print(' - Refreshing list.')

  global reprintListRef
  reprintListRef = []

  reprintsRef = db.collection('reprints')
  docs = reprintsRef.stream()

  for item in reprintList.get_children():
        reprintList.delete(item)

  newDocs = []

  for doc in docs:
    docObject = doc.to_dict()
    docObject.update({'id' : doc.id})

    if str(docObject['created']).__contains__('.') :
      docObject.update({'created' : int(str(docObject['created'])[:-7])})
    else : 
      docObject.update({'created' : int(str(docObject['created'])[:-3])})

    newDocs.append(docObject)

  newlist = sorted(newDocs, key=lambda x: x['created'], reverse=True)

  for docObject in newlist :

    if filters != [] :
      if not filters.__contains__(docObject['status']) :
        continue

    tempQty = f'{docObject["qtyNeeded"]} / {docObject["qtytotal"]}'

    timeStamp = datetime.datetime.fromtimestamp(docObject['created'])
    timeStamp = timeStamp.strftime("%m/%d/%Y, %H:%M:%S")

    sku = str(docObject['itemId'])

    values = [str(docObject['orderId']), str(docObject['itemId']), str(docObject['size']), str(tempQty), str(docObject['name']), str(docObject['reason']), str(timeStamp), str(docObject['status'])]

    reprintList.insert('', 'end',
        values = values, tags = (docObject['status']))

    reprintListRef.append(docObject)

#############################################

def pullFiles() :

  print(' - Pulling all requested reprints.')

  destFilePath = 'L:/Incoming Files/REPRINTS/'
  filesToGet = []

  # GET LIST OF REPRINTS
  reprintsRef = db.collection('reprints')
  docs = reprintsRef.stream()

  # FILTER TO ONLY DOCUMENTS THAT WERE REQUESTED
  for doc in docs:
    docObject = doc.to_dict()
    docObject.update({'id' : doc.id})

    if docObject['status'] == 'requested' :
      filesToGet.append(docObject)

  # FOR EACH FILE
  for i in filesToGet :
    
    orderNumber = i['orderId']
    sku = i['itemId']
    size = i['size']
    qtyNeeded = i['qtyNeeded']
    qtyTotal = i['qtytotal']

    if size.__contains__('X') :
      size = f'{size[0]}x{size[2]}'

    printFileName = f'{orderNumber}_{sku}_{size}_qty {qtyTotal}_PRINT.pdf'
    ticketFileName = f'{orderNumber}_{sku}_{size}_qty {qtyTotal}_TICKET.pdf'

    printFilePath = f'L:/Archive/TechStyles Archive/{orderNumber[0:2]}000 - {orderNumber[0:2]}999/{orderNumber}/{printFileName}'
    ticketFilePath = f'L:/Archive/TechStyles Archive/{orderNumber[0:2]}000 - {orderNumber[0:2]}999/{orderNumber}/{ticketFileName}'

    if os.path.exists(printFilePath) :

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

    else :

      if os.path.exists(f'L:/Archive/TechStyles Archive/{orderNumber[0:2]}000 - {orderNumber[0:2]}999/{orderNumber}/{orderNumber}_{sku}_{size[2]}x{size[0]}_qty {qtyTotal}_PRINT.pdf') :
        printFileName = f'{orderNumber}_{sku}_{size[2]}x{size[0]}_qty {qtyTotal}_PRINT.pdf'
        ticketFileName = f'{orderNumber}_{sku}_{size[2]}x{size[0]}_qty {qtyTotal}_TICKET.pdf'
        
        printFilePath = f'L:/Archive/TechStyles Archive/{orderNumber[0:2]}000 - {orderNumber[0:2]}999/{orderNumber}/{printFileName}'
        ticketFilePath = f'L:/Archive/TechStyles Archive/{orderNumber[0:2]}000 - {orderNumber[0:2]}999/{orderNumber}/{ticketFileName}'

        shutil.copy(printFilePath, destFilePath + printFileName)
        shutil.copy(ticketFilePath, destFilePath + ticketFileName)
      
        newPrintFile = f'{orderNumber}_{sku}_{size[2]}x{size[0]}_qty {qtyNeeded}_PRINT.pdf'
        newTicketFile = f'{orderNumber}_{sku}_{size[2]}x{size[0]}_qty {qtyNeeded}_TICKET.pdf'

        os.rename(destFilePath + printFileName, destFilePath + newPrintFile)
        os.rename(destFilePath + ticketFileName, destFilePath + newTicketFile)

        doc_ref = db.collection('reprints').document(i['id'])
        doc_ref.update({
          'status': 'pulled'
        })

      elif os.path.exists(f'L:/Archive/TechStyles Archive/{orderNumber[0:2]}000 - {orderNumber[0:2]}999/{orderNumber}/') :
        print(f' - Could not find item {sku} within order {orderNumber}. Make sure that the size and total quantities are both typed in correctly. Request another reprint with the correct information if needed.')
        doc_ref = db.collection('reprints').document(i['id'])
        doc_ref.update({
          'status': 'deleted'
        })

      else : 
        print(f' - Could not find order {orderNumber}. Marking the reprint as deleted.')
        doc_ref = db.collection('reprints').document(i['id'])
        doc_ref.update({
          'status': 'deleted'
        })

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

def openReprintsFolder() :

  os.startfile('L:/Incoming Files/REPRINTS/')
  print(' - Opening Reprints folder.')

def updateFilters() :

  print(' - Updating filters.')
  global filters
  filters = []

  if filter_Requested.get() == 1 :
    filters.append('requested')
  if filter_Pulled.get() == 1 :
    filters.append('pulled')
  if filter_Deleted.get() == 1 :
    filters.append('deleted')
  if filter_Printed.get() == 1 :
    filters.append('printed')
  if filter_Complete.get() == 1 :
    filters.append('complete')

  refresh()

##############################

def onResize(event) :

  if root.winfo_width() < 650 :
    requestButton.config(text = 'New')
    pullButton.config(text = 'Pull')
    statusButton.config(text = 'Status')
  else :
    requestButton.config(text = 'New Reprint')
    pullButton.config(text = 'Pull Requested Reprints')
    statusButton.config(text = 'Change Status')

def startApp() :

  print('Start')

  global root
  root = Tk()
  root.title('Citra Communications - Large Format Reprint Manager')
  root.config(bg = '#FFFFFF')
  root.minsize(width = 896, height = 400)
  root.bind("<Configure>", onResize)

  rootFrame = Frame(root, bg = '#FFFFFF')
  rootFrame.grid(padx = 32, pady = 32, sticky = 'NEWS')

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
  headerFrame.grid_columnconfigure(1, weight = 1)
  headerFrame.grid_columnconfigure(2, weight = 1000)
  headerFrame.grid_columnconfigure(3, weight = 1)
  headerFrame.grid_columnconfigure(4, weight = 1000)
  headerFrame.grid_columnconfigure(5, weight = 1)
  headerFrame.grid_columnconfigure(6, weight = 1000)
  headerFrame.grid_columnconfigure(7, weight = 1)
  headerFrame.grid_columnconfigure(8, weight = 100)
  headerFrame.grid_rowconfigure(0, weight = 1)
  headerFrame.grid_rowconfigure(1, weight = 1000)

  headerLabel = Label(headerFrame, text = 'Citra Reprint Manager', font = ('Segoe ui bold', 24), bg = 'white', justify = 'left')
  headerLabel.grid(column = 0, row = 0, columnspan = 4, sticky = 'NW', ipady = 8)

  global requestButton
  requestButton = Hoverbutton.HoverButton(headerFrame, root, text = 'New Reprint', command = newRequest, font = fontMain, colors = ['#158754', '#127448', '#156D44'], fg = 'white')
  requestButton.grid(column = 0, row = 1, ipadx = 16, ipady = 16, sticky = 'NEWS')

  seperatorFrame1 = Frame(headerFrame, bg = 'white', width = 16)
  seperatorFrame1.grid(column = 1, row = 1, sticky = 'NEWS')

  global pullButton
  pullButton = Hoverbutton.HoverButton(headerFrame, root, text = 'Pull Requested Reprints', command = pullFiles, font = fontMain)
  pullButton.grid(column = 2, row = 1, ipadx = 16, ipady = 16, sticky = 'NEWS')

  seperatorFrame2 = Frame(headerFrame, bg = 'white', width = 16)
  seperatorFrame2.grid(column = 3, row = 1, sticky = 'NEWS')

  global statusButton
  statusButton = Hoverbutton.HoverButton(headerFrame, root, text = 'Change Status', command = changeStatusDialogue, font = fontMain)
  statusButton.grid(column = 4, row = 1, ipadx = 16, ipady = 16, sticky = 'NEWS')

  seperatorFrame3 = Frame(headerFrame, bg = 'white', width = 16)
  seperatorFrame3.grid(column = 5, row = 1, sticky = 'NEWS')

  openReprintsFolderButton = Hoverbutton.HoverButton(headerFrame, root, text = 'Open Reprints Folder', command = openReprintsFolder, font = fontMain)
  openReprintsFolderButton.grid(column = 6, row = 1, ipadx = 16, ipady = 16, sticky = 'NEWS')

  seperatorFrame4 = Frame(headerFrame, bg = 'white', width = 16)
  seperatorFrame4.grid(column = 7, row = 1, sticky = 'NEWS')

  refreshButton = Hoverbutton.HoverButton(headerFrame, root, text = 'Refresh', command = refresh, font = fontMain)
  refreshButton.grid(column = 8, row = 1, ipadx = 16, ipady = 16, sticky = 'NEWS')

  ###

  # seperatorFrame = Frame(rootFrame, bg = '#E6E6E6', height = 2)
  # seperatorFrame.grid(column = 0, row = 1, sticky = 'NEWS', pady = 8)

  filterFrame = Frame(rootFrame, bg = 'white')
  filterFrame.grid(column = 0, row = 1, pady = 16, sticky = 'NWS')

  filterLabel = Label(filterFrame, bg = 'white', text = 'Filters:  ', font = fontMain)
  filterLabel.pack(side = 'left')

  global filter_Requested
  filter_Requested = IntVar()
  filter1 = Checkbutton(filterFrame, bg = 'white', text = 'Requested', command = updateFilters, variable = filter_Requested, font = fontMain)
  filter1.pack(side = 'left', padx = 8)

  global filter_Pulled
  filter_Pulled = IntVar()
  filter2 = Checkbutton(filterFrame, bg = 'white', text = 'Pulled', command = updateFilters, variable = filter_Pulled, font = fontMain)
  filter2.pack(side = 'left', padx = 8)

  global filter_Deleted
  filter_Deleted = IntVar()
  filter3 = Checkbutton(filterFrame, bg = 'white', text = 'Deleted', command = updateFilters, variable = filter_Deleted, font = fontMain)
  filter3.pack(side = 'left', padx = 8)

  global filter_Printed
  filter_Printed = IntVar()
  filter4 = Checkbutton(filterFrame, bg = 'white', text = 'Printed', command = updateFilters, variable = filter_Printed, font = fontMain)
  filter4.pack(side = 'left', padx = 8)

  global filter_Complete
  filter_Complete = IntVar()
  filter5 = Checkbutton(filterFrame, bg = 'white', text = 'Complete', command = updateFilters, variable = filter_Complete, font = fontMain)
  filter5.pack(side = 'left', padx = 8)

  ###

  listFrame = Frame(rootFrame, bg = '#FFFFFF')
  listFrame.grid(row = 2, column = 0, sticky = 'NEWS')

  listFrame.grid_columnconfigure(0, weight = 1000)
  listFrame.grid_columnconfigure(1, weight = 1)
  listFrame.grid_rowconfigure(0, weight = 1000)
  listFrame.grid_rowconfigure(1, weight = 1)

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

  reprintList.column('#0', width = 16, minwidth = 16, anchor = 'w', stretch = False)
  reprintList.column('Order', width = 128, minwidth = 96, anchor = 'w', stretch = False)
  reprintList.column('Sku', width = 128, minwidth = 96, anchor = 'w', stretch = False)
  reprintList.column('Size', width = 128, minwidth = 96, anchor = 'w', stretch = False)
  reprintList.column('Quantity', width = 128, minwidth = 96, anchor = 'w', stretch = False)
  reprintList.column('Submitter Name', width = 256, minwidth = 96, anchor = 'w')
  reprintList.column('Reason', width = 128, minwidth = 96, anchor = 'w')
  reprintList.column('Time', width = 256, minwidth = 128, anchor = 'w')
  reprintList.column('Status', width = 128, minwidth = 96, anchor = 'w', stretch = False)

  reprintList.tag_configure('deleted', background=colors['deleted'])
  reprintList.tag_configure('complete', background=colors['complete'])
  reprintList.tag_configure('requested', background=colors['requested'])
  reprintList.tag_configure('pulled', background=colors['pulled'])
  reprintList.tag_configure('printed', background=colors['printed'])

  scrollbar = Scrollbar(listFrame)
  scrollbar.grid(column = 1, row = 0, sticky = 'NEWS')

  reprintList.configure(yscrollcommand = scrollbar.set)
  scrollbar.config(command = reprintList.yview)
  # reprintList.bind('<Button-1>', selectItem)

  refresh()

  root.mainloop()

if __name__ == '__main__' :
  startApp()