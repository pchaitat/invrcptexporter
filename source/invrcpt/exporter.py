#!/usr/bin/env python3

import uno
from com.sun.star.beans import PropertyValue

def export_invoice_nowt_to_pdf(model: uno.pyuno, list_row: int,
  dest: 'path_to_pdf_file'):
  """
  export invoice with no withholding tax into pdf file

  parameters
    - model is spreadsheet file
  """

  # select the sheet to be printed
  inv_nowt_sheet = model.Sheets.getByName('inv-nowt')
  model.getCurrentController().setActiveSheet(inv_nowt_sheet)

  # have to select the area to be printed
  fdata = []
  fdata1 = PropertyValue()
  fdata1.Name = 'Selection'
  oCellRange = inv_nowt_sheet.getCellRangeByName('a1:h30')
  fdata1.Value = oCellRange
  fdata.append(fdata1)

  args = []
  arg1 = PropertyValue()
  arg1.Name = 'FilterName'
  arg1.Value = 'calc_pdf_Export'
  arg2 = PropertyValue()
  arg2.Name = 'FilterData'
  arg2.Value = uno.Any("[]com.sun.star.beans.PropertyValue", tuple(fdata))
  args.append(arg1)
  args.append(arg2)

  model.storeToURL('file:///' + dest, tuple(args))
