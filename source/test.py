#!/usr/bin/env python3

from invrcpt import exporter
from PyPDF2 import PdfFileReader

import invrcpt
import glob
import os
import shutil
import subprocess
import time
import unittest
import uno

class TestExportToPdf(unittest.TestCase):

  def setUp(self):
    BASE_DIR = os.environ['SCRIPT_PATH'] + '/'

    # create a libre calc file (from the sample file)
    self.dir_for_tmp_files_path = BASE_DIR + 'test.d/'
    self.invoice_list_file_path = self.dir_for_tmp_files_path + 'test.ods'
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

    # delete all pdf test files
    pdf_file_list = glob.glob(self.dir_for_tmp_files_path + '*.pdf')
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

  def extract_text(self, pdf_file_path: str) -> 'text':
    return subprocess.getoutput('pdftotext ' + pdf_file_path + ' -')

  def assert_tuple_of_str_in(self, str_tuple: tuple, text_to_check: str):
    for s in str_tuple:
      self.assertIn(s, text_to_check)

  def test_can_export_invoice_no_withholding_tax_into_pdf(self):
    list_row = 2 # specify the row in worksheet 'list'
    pdf_file_path = self.dir_for_tmp_files_path + 'invoice_no_wt.pdf'

    # set which invoice to be generated
    # inv-nowt stands for INVoice - NO Withholding Tax
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
    extracted_text = self.extract_text(pdf_file_path)

    self.assert_tuple_of_str_in((
      'Ms. Test Testing',
      '17/1 Testing Road, Testing Circle, Bangkok, 10100',
      '8. Apr. 2017',
      'KVM Based VirtualHosting Training and Tools Development',
      '441,000.00',
      'No.: 963'), extracted_text)

  def test_can_set_exported_invoice_pdf_file_with_correct_name(self):
    """
    wt stands for Withholding Tax
    """

    list_row = 2 # specify the row in worksheet 'list'

    invoice_filename = exporter.get_exported_invoice_pdf_filename(
      model=self.model, list_row=list_row)

    self.assertEqual(invoice_filename, 'invoice-963.pdf')

    list_row = 5

    invoice_filename = exporter.get_exported_invoice_pdf_filename(
      model=self.model, list_row=list_row)

    self.assertEqual(invoice_filename, 'invoice-966.pdf')

  def test_can_export_multiple_invoice_no_wt_into_pdf(self):
    """
    wt stands for Withholding Tax
    """

    # prepare data
    list_rows = (2, 3, 4) # specify the rows in worksheet 'list'
    exported_pdf_dir = self.dir_for_tmp_files_path

    # export!
    exporter.export_multiple_invoice_nowt_to_pdf(model=self.model,
      list_rows=list_rows, dest=exported_pdf_dir)

    # test

    # test if the pdf files exists
    self.assertTrue(os.path.isfile(exported_pdf_dir + 'invoice-963.pdf'))
    self.assertTrue(os.path.isfile(exported_pdf_dir + 'invoice-964.pdf'))
    self.assertTrue(os.path.isfile(exported_pdf_dir + 'invoice-965.pdf'))

    # test that each pdf file has 1 page
    with open(exported_pdf_dir + 'invoice-963.pdf', 'rb') as pdf_file:
      the_pdf_file = PdfFileReader(pdf_file)
      self.assertEqual(the_pdf_file.getNumPages(), 1)
    with open(exported_pdf_dir + 'invoice-964.pdf', 'rb') as pdf_file:
      the_pdf_file = PdfFileReader(pdf_file)
      self.assertEqual(the_pdf_file.getNumPages(), 1)
    with open(exported_pdf_dir + 'invoice-965.pdf', 'rb') as pdf_file:
      the_pdf_file = PdfFileReader(pdf_file)
      self.assertEqual(the_pdf_file.getNumPages(), 1)

    # test if each pdf file contains expected texts
    extracted_text = self.extract_text(exported_pdf_dir + 'invoice-963.pdf')
    self.assert_tuple_of_str_in((
      'Ms. Test Testing',
      '17/1 Testing Road, Testing Circle, Bangkok, 10100',
      '8. Apr. 2017',
      'KVM Based VirtualHosting Training and Tools Development',
      '441,000.00',
      'No.: 963'), extracted_text)

    extracted_text = self.extract_text(exported_pdf_dir + 'invoice-964.pdf')
    self.assert_tuple_of_str_in((
      'Mr. Testa Testo',
      '17/1 Testing Road, Testing Sqaure, Bangkok, 10100',
      '8. Apr. 2017',
      'Emperor + uWSGI system development',
      '16,000.00',
      'No.: 964'), extracted_text)

    extracted_text = self.extract_text(exported_pdf_dir + 'invoice-965.pdf')
    self.assert_tuple_of_str_in((
      'Ms. Testest Threst',
      '17/1 Testing Road, Testing Triangle, Bangkok, 10100',
      '8. Apr. 2017',
      'Banner graphics design',
      '6,000.00',
      'No.: 965'), extracted_text)

    self.fail('Finish the test!')

if __name__ == '__main__':
  unittest.main()
