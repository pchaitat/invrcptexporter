#!/usr/bin/env python3

from invrcpt import exporter
from PyPDF2 import PdfFileReader

import invrcpt
import glob
import os
import shutil
import socket
import subprocess
import time
import unittest
import uno

class TestExportToPdf(unittest.TestCase):

  def setUp(self):
    BASE_DIR = os.environ['SCRIPT_PATH'] + '/'

    # create a libre calc file (from the sample file)
    self.invoice_list_file_path = BASE_DIR + 'test.d/test.ods'
    shutil.copy2(BASE_DIR + 'sample_invoice_list.ods',
      self.invoice_list_file_path)

    subprocess.Popen("libreoffice --accept='socket,host=localhost," +
      "port=2002;urp;StarOffice.ServiceManager' " +
      self.invoice_list_file_path, shell=True)

    # it takes some time for libreoffice program to start, if we don't
    # wait, we can't connect to its socket
    time.sleep(2)

    self.initialize_uno()

  def tearDown(self):
    self.model.close(True)
    # delete the libre calc test files
    os.remove(self.invoice_list_file_path)
    pdf_file_list = glob.glob(os.environ['SCRIPT_PATH'] +
      '/test.d/*.pdf')
    for f in pdf_file_list:
      os.remove(f)

    time.sleep(2)

  def initialize_uno(self):
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext(
      "com.sun.star.bridge.UnoUrlResolver", localContext)
    context = resolver.resolve(
      "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    desktop = context.ServiceManager.createInstanceWithContext(
      "com.sun.star.frame.Desktop", context)
    self.model = desktop.getCurrentComponent()

    time.sleep(2)

  def test_can_export_invoice_into_pdf(self):
    list_row = 2
    pdf_file_path = os.environ['SCRIPT_PATH'] + '/test.d/invoice_no_wt.pdf'

    # select which invoice to be generated
    inv_nowt_sheet = self.model.Sheets.getByName('inv-nowt')
    list_row_cell = inv_nowt_sheet.getCellRangeByName('h1')
    list_row_cell.Value = list_row

    # export it using our export function
    exporter.export_invoice_nowt_to_pdf(model=self.model,
      list_row=list_row, dest=pdf_file_path)

    # test that the generated pdf matches what we expect

    # test if the pdf file exists
    self.assertTrue(os.path.isfile(pdf_file_path))

    # test that the pdf file has 1 page
    with open(pdf_file_path, 'rb') as pdf_file:
      the_pdf_file = PdfFileReader(pdf_file)
      self.assertEqual(the_pdf_file.getNumPages(), 1)

    # test if the pdf contains expected texts
    self.fail('Finish the test!')

if __name__ == '__main__':
  unittest.main()
