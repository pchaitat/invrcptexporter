#!/usr/bin/env python3

# we have to import uno before PropertyValue for some reason I don't
# know yet, the result is that we can't sort the import statements
# as usual
import uno

from com.sun.star.beans import PropertyValue
import time

# common form name list
INV_NOWT_FORM  = 'inv-nowt'
INV_WT3_FORM   = 'inv-wt3'
RCPT_NOWT_FORM = 'rcpt-nowt'
RCPT_WT3_FORM  = 'rcpt-wt3'

class Ods:

  def __init__(self):

    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext(
      "com.sun.star.bridge.UnoUrlResolver", localContext)
    context = resolver.resolve(
      "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    desktop = context.ServiceManager.createInstanceWithContext(
      "com.sun.star.frame.Desktop", context)
    self.model = desktop.getCurrentComponent()

    time.sleep(2)

  def get_exported_invoice_pdf_filename(self, list_row: int):
    # get invoice number
    list_sheet = self.model.Sheets.getByName('list')
    invoice_number = list_sheet.getCellRangeByName('b' + str(list_row)).Value

    return 'invoice-' + ("%.10g" % invoice_number) + '.pdf'

  def get_exported_receipt_pdf_filename(self, list_row: int):
    # get receipt number
    list_sheet = self.model.Sheets.getByName('list')
    receipt_number = list_sheet.getCellRangeByName('h' + str(list_row)).Value

    return 'receipt-' + ("%.10g" % receipt_number) + '.pdf'

  def export_active_sheet_to_pdf(self, dest: 'path_to_pdf_file'):
    """notes: assuming the area of the active sheet to be printed is
    'a1:h30'
    """

    active_sheet = self.model.getCurrentController().getActiveSheet()

    # have to select the area to be printed
    fdata = []
    fdata1 = PropertyValue()
    fdata1.Name = 'Selection'
    oCellRange = active_sheet.getCellRangeByName('a1:h30')
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

    self.model.storeToURL('file:///' + dest, tuple(args))

  def export_multiple_common_form_to_pdf(self, form: 'form_name',
    list_rows: tuple, dest: 'exported_pdf_dir'):
    """
    keyword arguments:
      form_name -- the name of the form which is the same name as the
                   name of the worksheet to be printed which is defined
                   above, see "common form name list" at the top of this
                   file (INV_NOWT_FORM, INV_WT3_FORM, etc)
      list_rows -- tuple of rows in list worksheet, e.g. (2, 3, 5, 6)
      dest      -- directory to export pdf files into
    """

    # select the sheet to be printed
    the_sheet = self.model.Sheets.getByName(form)
    self.model.getCurrentController().setActiveSheet(the_sheet)

    for list_row in list_rows:
      the_sheet.getCellRangeByName('h1').Value = list_row

      if form in (INV_NOWT_FORM, INV_WT3_FORM):
        exported_pdf_filename = \
          self.get_exported_invoice_pdf_filename(list_row)
      else:
        exported_pdf_filename = \
          self.get_exported_receipt_pdf_filename(list_row)

      self.export_active_sheet_to_pdf(''.join((dest,
        exported_pdf_filename)))
