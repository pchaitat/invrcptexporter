#!/usr/bin/env python3

import uno
from com.sun.star.beans import PropertyValue

def get_exported_invoice_pdf_filename(model: uno.pyuno, list_row: int):
  # get invoice number
  list_sheet = model.Sheets.getByName('list')
  invoice_number = list_sheet.getCellRangeByName('b' + str(list_row)).Value

  return 'invoice-' + ("%.10g" % invoice_number) + '.pdf'

def export_invoice_nowt_to_pdf(model: uno.pyuno, list_row: int,
  dest: 'path_to_pdf_file'):
  """
  export invoice with no withholding tax into pdf file

  parameters
    - model: the spreadsheet file
  """

  # select the sheet to be printed
  inv_nowt_sheet = model.Sheets.getByName('inv-nowt')
  model.getCurrentController().setActiveSheet(inv_nowt_sheet)
  inv_nowt_sheet.getCellRangeByName('h1').Value = list_row

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

def export_multiple_invoice_nowt_to_pdf(model: uno.pyuno,
  list_rows: tuple, dest: 'exported_pdf_dir'):
  """
  keyword arguments:
    model     -- the spreadsheet file
    list_rows -- tuple of rows in list worksheet
    dest      -- directory to export pdf files into
  """

  for list_row in list_rows:
    export_invoice_nowt_to_pdf(
      model=model,
      list_row=list_row,
      dest=dest + get_exported_invoice_pdf_filename(model, list_row))
