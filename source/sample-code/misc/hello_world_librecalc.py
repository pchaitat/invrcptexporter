#!/usr/bin/env python3

# export_to_pdf.py
#
# references
#   - https://onesheep.org/scripting-libreoffice-python/
#   - http://christopher5106.github.io/office/2015/12/06/openoffice-lib
#     reoffice-automate-your-office-tasks-with-python-macros.html

import uno
from com.sun.star.beans import PropertyValue

localContext = uno.getComponentContext()
resolver = localContext.ServiceManager.createInstanceWithContext(
  "com.sun.star.bridge.UnoUrlResolver", localContext)
context = resolver.resolve(
  "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
desktop = context.ServiceManager.createInstanceWithContext(
  "com.sun.star.frame.Desktop", context)
model = desktop.getCurrentComponent()

# pdf
#properties=[]
#p=PropertyValue()
#p.Name='FilterName'
#p.Value='calc_pdf_Export'
#properties.append(p)
#model.storeToURL('file:///home/ubuntu/tmp/testout.pdf',tuple(properties))


model.close(True)

# access the active sheet
active_sheet = model.CurrentController.ActiveSheet

# access cell C4
cell1 = active_sheet.getCellRangeByName("C4")

# set text inside
cell1.String = "Hello world"

# other example with a value
cell2 = active_sheet.getCellRangeByName("E6")
cell2.Value = cell2.Value + 1

import sys
sys.exit(0)


