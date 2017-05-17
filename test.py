#!/usr/bin/env python3

from invrcpt.exporter import Ods, INV_NOWT_FORM, INV_WT3_FORM, \
  RCPT_NOWT_FORM, RCPT_WT3_FORM
from PyPDF2 import PdfFileReader

import invrcpt
import glob
import os
import re
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

    self.ods = Ods()

  def tearDown(self):
    self.ods.model.close(True)

    # delete the libre calc test files
    os.remove(self.invoice_list_file_path)

    # delete all pdf test files
    pdf_file_list = glob.glob(self.dir_for_tmp_files_path + '*.pdf')
    for f in pdf_file_list:
      os.remove(f)

    time.sleep(2)

  def get_row_dict_from_list_sheet(self, list_row: int):
    """
    go to 'list' worksheet, go to row, 'list_row', get data from each
    column into tuple and return it
    """

    list_sheet = self.ods.model.Sheets.getByName('list')
    row_dict = {}
    row_dict['invoice_date'] = list_sheet.getCellRangeByName(
      'a'+str(list_row)).String
    row_dict['invoice_id'] = list_sheet.getCellRangeByName(
      'b'+str(list_row)).String
    row_dict['price'] = list_sheet.getCellRangeByName(
      'c'+str(list_row)).String[1:]
    row_dict['wt'] = list_sheet.getCellRangeByName(
      'e'+str(list_row)).String[1:]
    row_dict['receipt_date'] = list_sheet.getCellRangeByName(
      'g'+str(list_row)).String
    row_dict['receipt_id'] = list_sheet.getCellRangeByName(
      'h'+str(list_row)).String
    row_dict['customer_info_0'] = list_sheet.getCellRangeByName(
      'i'+str(list_row)).String
    row_dict['customer_info_1'] = list_sheet.getCellRangeByName(
      'j'+str(list_row)).String
    row_dict['customer_info_2'] = list_sheet.getCellRangeByName(
      'k'+str(list_row)).String
    row_dict['project_name'] = list_sheet.getCellRangeByName(
      'm'+str(list_row)).String

    return row_dict

  def extract_text(self, pdf_file_path: str) -> 'text':
    return re.sub('\\n', ' ',
      subprocess.getoutput('pdftotext ' + pdf_file_path + ' -'))

  def assert_tuple_of_str_in(self, str_tuple: tuple, text_to_check: str):
    for s in str_tuple:
      self.assertIn(s, text_to_check)

  def test_can_set_exported_invoice_pdf_file_with_correct_name(self):
    """
    wt stands for Withholding Tax
    """

    list_row = 2 # specify the row in worksheet 'list'

    invoice_filename = self.ods.get_exported_invoice_pdf_filename(
      list_row=list_row)

    self.assertEqual(invoice_filename, 'invoice-963.pdf')

    list_row = 5

    invoice_filename = self.ods.get_exported_invoice_pdf_filename(
      list_row=list_row)

    self.assertEqual(invoice_filename, 'invoice-966.pdf')

  def test_can_set_exported_receipt_pdf_file_with_correct_name(self):
    """
    wt stands for Withholding Tax
    """

    list_row = 2 # specify the row in worksheet 'list'

    receipt_filename = self.ods.get_exported_receipt_pdf_filename(
      list_row=list_row)

    self.assertEqual(receipt_filename, 'receipt-800.pdf')

    list_row = 5

    receipt_filename = self.ods.get_exported_receipt_pdf_filename(
      list_row=list_row)

    self.assertEqual(receipt_filename, 'receipt-925.pdf')

  def test_can_export_multiple_invoice_no_wt_into_pdf(self):
    """
    wt stands for Withholding Tax
    """

    # prepare data
    list_rows = (2, 3, 4) # specify the rows in worksheet 'list'
    exported_pdf_dir = self.dir_for_tmp_files_path

    # export!
    self.ods.export_multiple_common_form_to_pdf(form=INV_NOWT_FORM,
      list_rows=list_rows, dest=exported_pdf_dir)

    # test

    for list_row in list_rows:
      filename = self.ods.get_exported_invoice_pdf_filename(list_row)

      # test if the pdf files exists
      self.assertTrue(os.path.isfile(exported_pdf_dir+filename))

      # test that each pdf file has 1 page
      with open(exported_pdf_dir + filename, 'rb') as pdf_file:
        the_pdf_file = PdfFileReader(pdf_file)
        self.assertEqual(the_pdf_file.getNumPages(), 1)

      # test if each pdf file contains expected texts
      extracted_text = self.extract_text(exported_pdf_dir+filename)

      expected_text_dict = self.get_row_dict_from_list_sheet(list_row)
      expected_text_tuple = (
        expected_text_dict['invoice_date'],
        expected_text_dict['invoice_id'],
        expected_text_dict['price'],
        expected_text_dict['customer_info_0'],
        expected_text_dict['customer_info_1'],
        expected_text_dict['customer_info_2'],
        expected_text_dict['project_name'],)

      self.assert_tuple_of_str_in(expected_text_tuple, extracted_text)

  def test_can_export_multiple_invoice_wt3_into_pdf(self):
    """
    wt stands for Withholding Tax
    """

    # prepare data
    list_rows = (5, 6) # specify the rows in worksheet 'list'
    exported_pdf_dir = self.dir_for_tmp_files_path

    # export!
    self.ods.export_multiple_common_form_to_pdf(form=INV_WT3_FORM,
      list_rows=list_rows, dest=exported_pdf_dir)

    # test

    for list_row in list_rows:
      filename = self.ods.get_exported_invoice_pdf_filename(list_row)

      # test if the pdf files exists
      self.assertTrue(os.path.isfile(exported_pdf_dir+filename))

      # test that each pdf file has 1 page
      with open(exported_pdf_dir + filename, 'rb') as pdf_file:
        the_pdf_file = PdfFileReader(pdf_file)
        self.assertEqual(the_pdf_file.getNumPages(), 1)

      # test if each pdf file contains expected texts
      extracted_text = self.extract_text(exported_pdf_dir+filename)

      expected_text_dict = self.get_row_dict_from_list_sheet(list_row)
      expected_text_tuple = (
        expected_text_dict['invoice_date'],
        expected_text_dict['invoice_id'],
        expected_text_dict['price'],
        expected_text_dict['customer_info_0'],
        expected_text_dict['customer_info_1'],
        expected_text_dict['customer_info_2'],
        expected_text_dict['project_name'],
        expected_text_dict['wt'],)

      self.assert_tuple_of_str_in(expected_text_tuple, extracted_text)

  def test_can_export_multiple_receipt_no_wt_into_pdf(self):
    """
    wt stands for Withholding Tax
    """

    # prepare data
    list_rows = (2, 3, 4) # specify the rows in worksheet 'list'
    exported_pdf_dir = self.dir_for_tmp_files_path

    # export!
    self.ods.export_multiple_common_form_to_pdf(form=RCPT_NOWT_FORM,
      list_rows=list_rows, dest=exported_pdf_dir)

    # test

    for list_row in list_rows:
      filename = self.ods.get_exported_receipt_pdf_filename(list_row)

      # test if the pdf files exists
      self.assertTrue(os.path.isfile(exported_pdf_dir+filename))

      # test that each pdf file has 1 page
      with open(exported_pdf_dir + filename, 'rb') as pdf_file:
        the_pdf_file = PdfFileReader(pdf_file)
        self.assertEqual(the_pdf_file.getNumPages(), 1)

      # test if each pdf file contains expected texts
      extracted_text = self.extract_text(exported_pdf_dir+filename)

      expected_text_dict = self.get_row_dict_from_list_sheet(list_row)
      expected_text_tuple = (
        expected_text_dict['invoice_id'],
        expected_text_dict['price'],
        expected_text_dict['receipt_date'],
        expected_text_dict['receipt_id'],
        expected_text_dict['customer_info_0'],
        expected_text_dict['customer_info_1'],
        expected_text_dict['customer_info_2'],
        expected_text_dict['project_name'],)

      self.assert_tuple_of_str_in(expected_text_tuple, extracted_text)

  def test_can_export_multiple_receipt_wt3_into_pdf(self):
    """
    wt stands for Withholding Tax
    """

    # prepare data
    list_rows = (5, 6) # specify the rows in worksheet 'list'
    exported_pdf_dir = self.dir_for_tmp_files_path

    # export!
    self.ods.export_multiple_common_form_to_pdf(form=RCPT_WT3_FORM,
      list_rows=list_rows, dest=exported_pdf_dir)

    # test

    for list_row in list_rows:
      filename = self.ods.get_exported_receipt_pdf_filename(list_row)

      # test if the pdf files exists
      self.assertTrue(os.path.isfile(exported_pdf_dir+filename))

      # test that each pdf file has 1 page
      with open(exported_pdf_dir + filename, 'rb') as pdf_file:
        the_pdf_file = PdfFileReader(pdf_file)
        self.assertEqual(the_pdf_file.getNumPages(), 1)

      # test if each pdf file contains expected texts
      extracted_text = self.extract_text(exported_pdf_dir+filename)

      expected_text_dict = self.get_row_dict_from_list_sheet(list_row)
      expected_text_tuple = (
        expected_text_dict['invoice_id'],
        expected_text_dict['price'],
        expected_text_dict['receipt_date'],
        expected_text_dict['receipt_id'],
        expected_text_dict['customer_info_0'],
        expected_text_dict['customer_info_1'],
        expected_text_dict['customer_info_2'],
        expected_text_dict['project_name'],
        expected_text_dict['wt'],)

      self.assert_tuple_of_str_in(expected_text_tuple, extracted_text)
    self.fail('Finish the test!')

if __name__ == '__main__':
  unittest.main()
