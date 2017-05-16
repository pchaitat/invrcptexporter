#!/usr/bin/env python3

from invrcpt.exporter import Ods
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

  def get_data_tuple_from_list_sheet(self, list_row: int):
    """
    go to 'list' worksheet, go to row, 'list_row', get data from each
    column into tuple and return it
    """

    list_sheet = self.ods.model.Sheets.getByName('list')
    invoice_date = list_sheet.getCellRangeByName('a'+str(list_row)).String
    invoice_id = list_sheet.getCellRangeByName('b'+str(list_row)).String
    price = list_sheet.getCellRangeByName('c'+str(list_row)).String[1:]
    customer_info_0 = list_sheet.getCellRangeByName('i'+str(list_row)).String
    customer_info_1 = list_sheet.getCellRangeByName('j'+str(list_row)).String
    customer_info_2 = list_sheet.getCellRangeByName('k'+str(list_row)).String
    project_name = list_sheet.getCellRangeByName('m'+str(list_row)).String

    return (invoice_date, invoice_id, price, customer_info_0,
      customer_info_1, customer_info_2, project_name)

  def extract_text(self, pdf_file_path: str) -> 'text':
    return subprocess.getoutput('pdftotext ' + pdf_file_path + ' -')

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

  def test_can_export_invoice_no_withholding_tax_into_pdf(self):
    list_row = 2 # specify the row in worksheet 'list'
    pdf_file_path = ''.join((self.dir_for_tmp_files_path,
      self.ods.get_exported_invoice_pdf_filename(list_row)))

    # export it using our export function
    self.ods.export_invoice_nowt_to_pdf(list_row=list_row, dest=pdf_file_path)

    # test that the generated pdf matches what we expect

    # test if the pdf file exists
    self.assertTrue(os.path.isfile(pdf_file_path))

    # test that the pdf file has 1 page
    with open(pdf_file_path, 'rb') as pdf_file:
      the_pdf_file = PdfFileReader(pdf_file)
      self.assertEqual(the_pdf_file.getNumPages(), 1)

    # test if the pdf contains expected texts
    extracted_text = self.extract_text(pdf_file_path)

    self.assert_tuple_of_str_in(
      self.get_data_tuple_from_list_sheet(list_row),
      extracted_text)

  def test_can_export_multiple_invoice_no_wt_into_pdf(self):
    """
    wt stands for Withholding Tax
    """

    # prepare data
    list_rows = (2, 3, 4) # specify the rows in worksheet 'list'
    exported_pdf_dir = self.dir_for_tmp_files_path

    # export!
    self.ods.export_multiple_invoice_nowt_to_pdf(list_rows=list_rows,
      dest=exported_pdf_dir)

    # test

    file_name_0 = self.ods.get_exported_invoice_pdf_filename(list_rows[0])
    file_name_1 = self.ods.get_exported_invoice_pdf_filename(list_rows[1])
    file_name_2 = self.ods.get_exported_invoice_pdf_filename(list_rows[2])

    # test if the pdf files exists
    self.assertTrue(os.path.isfile(exported_pdf_dir+file_name_0))
    self.assertTrue(os.path.isfile(exported_pdf_dir+file_name_1))
    self.assertTrue(os.path.isfile(exported_pdf_dir+file_name_2))

    # test that each pdf file has 1 page
    with open(exported_pdf_dir + file_name_0, 'rb') as pdf_file:
      the_pdf_file = PdfFileReader(pdf_file)
      self.assertEqual(the_pdf_file.getNumPages(), 1)
    with open(exported_pdf_dir + file_name_1, 'rb') as pdf_file:
      the_pdf_file = PdfFileReader(pdf_file)
      self.assertEqual(the_pdf_file.getNumPages(), 1)
    with open(exported_pdf_dir + file_name_2, 'rb') as pdf_file:
      the_pdf_file = PdfFileReader(pdf_file)
      self.assertEqual(the_pdf_file.getNumPages(), 1)

    # test if each pdf file contains expected texts
    extracted_text = self.extract_text(exported_pdf_dir + file_name_0)
    self.assert_tuple_of_str_in(
      self.get_data_tuple_from_list_sheet(list_rows[0]),
      extracted_text)

    extracted_text = self.extract_text(exported_pdf_dir + file_name_1)
    self.assert_tuple_of_str_in(
      self.get_data_tuple_from_list_sheet(list_rows[1]),
      extracted_text)

    extracted_text = self.extract_text(exported_pdf_dir + file_name_2)
    self.assert_tuple_of_str_in(
      self.get_data_tuple_from_list_sheet(list_rows[2]),
      extracted_text)

    self.fail('Finish the test!')

if __name__ == '__main__':
  unittest.main()
