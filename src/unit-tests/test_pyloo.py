# coding: utf-8

"""
PyLOO - Python Libre/Open Office interface API (UNO)

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
import pyloo

###############################################################################
HIDE_OFFICE_RESULTS = True

###############################################################################


def pyloo_open_close_doc(f):
    """
    Decorator. Opens libre/open office document call function for testing
    and closes the document.
    """
    def _f(*args, **kwargs):
        # open document
        doc = pyloo.Document()
        file_name = os.getcwd() + "/test.ods"
        doc.open_document(file_name)

        # Test function call
        kwargs['doc'] = doc
        keep_doc_opened = _f._retval = f(*args, **kwargs)
        keep_doc_opened = bool(keep_doc_opened)

        # close document
        if HIDE_OFFICE_RESULTS and not keep_doc_opened:
            doc.close_document()
    return _f

###############################################################################


class Test_PyLOO_Document(unittest.TestCase):

    def setUp(self):
        pass

    def test_document_initialize(self):
        doc = pyloo.Document()
        self.assertFalse(doc.is_null())

    def test_document_new_save_close(self):
        doc = pyloo.Document()
        file_name_saved = os.getcwd() + "/test_saved.ods"
        self.assertTrue(doc.new_document())
        self.assertTrue(doc.save_document(file_name_saved))
        self.assertTrue(os.path.isfile(file_name_saved),
                        "File does not exists")
        # remove test file
        os.remove(file_name_saved)
        self.assertTrue(doc.close_document())

    def test_document_open_close(self):
        doc = pyloo.Document()
        file_name = os.getcwd() + "/test.ods"
        self.assertTrue(doc.open_document(file_name))
        self.assertTrue(doc.close_document())

    @pyloo_open_close_doc
    def test_document_sheets(self, doc):
        self.assertFalse(doc.sheets().is_null(), "get sheets object")

    @pyloo_open_close_doc
    def test_document_fields(self, doc):
        self.assertFalse(doc.fields().is_null(), "get fields object")

###############################################################################


class Test_PyLOO_Fields(unittest.TestCase):

    @pyloo_open_close_doc
    def test_fields_field(self, doc):
        field = doc.fields().field("TABLE_NAME")
        self.assertFalse(field.is_null(), "get field object")

    @pyloo_open_close_doc
    def test_fields_count(self, doc):
        self.assertEqual(doc.fields().count(), 11, "Wrong number of fields")

###############################################################################


class Test_PyLOO_Field(unittest.TestCase):

    @pyloo_open_close_doc
    def test_field_set_get(self, doc):
        field = doc.fields().field("TABLE_NAME")
        test_value = "Test table name"

        # set and get value without offset
        self.assertTrue(field.set_value(test_value))
        self.assertEqual(field.value(), test_value)

        # set and get value with offset
        self.assertTrue(field.set_value(test_value, 2, 1))
        self.assertEqual(field.value(2, 1), test_value)

    @pyloo_open_close_doc
    def test_field_insert_rows(self, doc):
        t1_field = doc.fields().field("FIELD_1")
        t2_field = doc.fields().field("T2FIELD_1")

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


class Test_PyLOO_Sheets(unittest.TestCase):

    @pyloo_open_close_doc
    def test_sheets_sheet_by_index(self, doc):
        sheet = doc.sheets().sheet("Sheet1")
        self.assertFalse(sheet.is_null(), "get sheet object")

    @pyloo_open_close_doc
    def test_sheets_insert_remove_spreadsheet_count(self, doc):
        self.assertTrue(doc.sheets().insert_spreadsheet("test1", 1))
        self.assertEqual(doc.sheets().count(), 2, "Wrong number of fields")
        self.assertTrue(doc.sheets().remove_spreadsheet("test1"))
        self.assertEqual(doc.sheets().count(), 1, "Wrong number of fields")

###############################################################################


class Test_PyLOO_Sheet(unittest.TestCase):

    @pyloo_open_close_doc
    def test_sheet_set_get_cell_value_by_index(self, doc):
        s_val = "value"
        n_val = 123
        f_val = 1.23
#         formula = "=G2+G3"

        sheet = doc.sheets().sheet("Sheet1")

        # set values
        self.assertTrue(sheet.set_cell_value_by_index(s_val, 7, 0))
        self.assertTrue(sheet.set_cell_value_by_index(n_val, 7, 1))
        self.assertTrue(sheet.set_cell_value_by_index(f_val, 7, 2))
        self.assertTrue(sheet.set_cell_value_by_index(formula, 7, 3, True))

        # get values and check results
        self.assertEqual(sheet.cell_value_by_index(7, 0), s_val)
        self.assertEqual(sheet.cell_value_by_index(7, 1), n_val)
        self.assertEqual(sheet.cell_value_by_index(7, 2), f_val)
#         self.assertEqual(sheet.set_cell_value_by_index(7, 3), formula))

###############################################################################


if __name__ == "__main__":
    unittest.main()
