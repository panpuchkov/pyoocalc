# coding: utf-8

"""
PyOOCalc - Python Libre/Open Office Calc interface API (UNO)

Requirements for Ubuntu users:

sudo apt-get install python-uno

Optional requirements:

sudo apt-get install libreoffice-script-provider-python

Copyright (c) 2015

@author: Yurii Puchkov
@organization: http://arilot.com/
@license: GPL v3
@contact: panpuchkov@gmail.com
"""

import unittest

###############################################################################
import os
import sys

sys.path.append('./../')
import pyoocalc

###############################################################################
HIDE_OFFICE_RESULTS = True

###############################################################################


def pyoocalc_open_close_doc(f):
    """
    Decorator. Opens libre/open office document call function for testing
    and closes the document.
    """
    def _f(*args, **kwargs):
        # open document
        doc = pyoocalc.Document()
        file_name = os.getcwd() + "/test.ods"
        doc.open_document(file_name)

        # Test function call
        kwargs['doc'] = doc
        keep_doc_opened = _f._retval = f(*args, **kwargs)
        keep_doc_opened = bool(keep_doc_opened)

        # close document
        if HIDE_OFFICE_RESULTS and not keep_doc_opened:
            doc.close_document()
        del doc
    return _f

###############################################################################


class Test_PyOOCalc_Document(unittest.TestCase):

    def setUp(self):
        pass

    def test_document_initialize(self):
        doc = pyoocalc.Document()
        self.assertFalse(doc.is_null)

    def test_document_new_save_close(self):
        doc = pyoocalc.Document()
        file_name_saved = os.getcwd() + "/test_saved.ods"
        self.assertTrue(doc.new_document())
        self.assertTrue(doc.save_document(file_name_saved))
        self.assertTrue(os.path.isfile(file_name_saved),
                        "File does not exists")
        # remove test file
        os.remove(file_name_saved)
        self.assertTrue(doc.close_document())

    def test_document_open_close(self):
        doc = pyoocalc.Document()
        file_name = os.getcwd() + "/test.ods"
        self.assertTrue(doc.open_document(file_name))
        self.assertTrue(doc.close_document())

    @pyoocalc_open_close_doc
    def test_document_sheets(self, doc):
        self.assertFalse(doc.sheets.is_null, "get sheets object")

    @pyoocalc_open_close_doc
    def test_document_fields(self, doc):
        self.assertFalse(doc.fields.is_null, "get fields object")

###############################################################################


class Test_PyOOCalc_Base(unittest.TestCase):
    """
    Setup base class for future tests.
    It opens and close the 'ods' document for testing.
    """
    def setUp(self):
        # open document
        self._doc = pyoocalc.Document()
        file_name = os.getcwd() + "/test.ods"
        self._doc.open_document(file_name)

    def tearDown(self):
        # close document
        self._doc.close_document()
        del self._doc

###############################################################################


class Test_PyOOCalc_Fields(Test_PyOOCalc_Base):

    def test_fields_field(self):
        field = self._doc.fields.field("TABLE_NAME")
        self.assertFalse(field.is_null, "get field object")

    def test_fields_count(self):
        self.assertEqual(self._doc.fields.count, 11,
                         "Wrong number of fields")

###############################################################################


class Test_PyOOCalc_Field(Test_PyOOCalc_Base):

    def test_field_set_get(self):
        field = self._doc.fields.field("TABLE_NAME")
        test_value = "Test table name"

        # set and get value without offset
        self.assertTrue(field.set_value(test_value))
        self.assertEqual(field.value(), test_value)

        # set and get value with offset
        self.assertTrue(field.set_value(test_value, 2, 1))
        self.assertEqual(field.value(2, 1), test_value)

    def test_field_insert_rows(self):
        t1_field = self._doc.fields.field("FIELD_1")
        t2_field = self._doc.fields.field("T2FIELD_1")

        def check_insert_rows(field, test_value, step):
            # insert row with step = `step`
            num_rows = 1
            self.assertTrue(field.insert_rows(num_rows=num_rows, step=step))

            # set value at row 2 with considering step = `step`
            self.assertTrue(field.set_value(test_value, 0, 1 + step))
            self.assertEqual(field.value(0, 1 + step), test_value)

            # insert two more rows with step = `step`
            num_rows = 2
            self.assertTrue(field.insert_rows(num_rows=num_rows, step=step))

            # get result at row 3 with considering step step = `step`
            self.assertEqual(field.value(0, 1 + ((num_rows + 1) * step)),
                             test_value)

        # check row insertion with row step 1 and row step 2
        check_insert_rows(t2_field, "t2.f1.1", 1)
        check_insert_rows(t1_field, "f1.1", 2)

###############################################################################


class Test_PyOOCalc_Sheets(Test_PyOOCalc_Base):

    def test_sheets_sheet_by_index(self):
        sheet = self._doc.sheets.sheet("Sheet1")
        self.assertFalse(sheet.is_null, "get sheet object")

    def test_sheets_insert_remove_spreadsheet_count(self):
        self.assertTrue(self._doc.sheets.insert_spreadsheet("test1", 1))
        self.assertEqual(self._doc.sheets.count, 2,
                         "Wrong number of fields")
        self.assertTrue(self._doc.sheets.remove_spreadsheet("test1"))
        self.assertEqual(self._doc.sheets.count, 1,
                         "Wrong number of fields")

###############################################################################


class Test_PyOOCalc_Sheet(Test_PyOOCalc_Base):

    def test_sheet_set_get_cell_value_by_index(self):
        s_val = "value"
        n_val = 123
        f_val = 1.23
#         formula = "=G2+G3"

        sheet = self._doc.sheets.sheet("Sheet1")

        # set values
        self.assertTrue(sheet.set_cell_value_by_index(s_val, 7, 0))
        self.assertTrue(sheet.set_cell_value_by_index(n_val, 7, 1))
        self.assertTrue(sheet.set_cell_value_by_index(f_val, 7, 2))
#         self.assertTrue(sheet.set_cell_value_by_index(formula, 7, 3, True))

        # get values and check results
        self.assertEqual(sheet.cell_value_by_index(7, 0), s_val)
        self.assertEqual(sheet.cell_value_by_index(7, 1), n_val)
        self.assertEqual(sheet.cell_value_by_index(7, 2), f_val)
#         self.assertEqual(sheet.set_cell_value_by_index(7, 3), formula))

###############################################################################


if __name__ == "__main__":
    unittest.main()
#     suite = unittest.TestLoader().loadTestsFromTestCase(Test_PyOOCalc_Sheet)
#     unittest.TextTestRunner(verbosity=2).run(suite)
