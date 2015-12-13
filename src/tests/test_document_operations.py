# coding: utf-8

import unittest
###############################################################################
import os
import sys
sys.path.append('./../')

###############################################################################
import pyloo

###############################################################################


def pyloo_open_close_doc(f):
    """
    Decorator. Opens libre/open office document call function for testing
    and closes the document.
    """

    def _f(*args, **kwargs):
        doc = pyloo.Document()
        file_name = os.getcwd() + "/test.ods"
        doc.open_document(file_name)

        kwargs['doc'] = doc
        _f._retval = f(*args, **kwargs)

        doc.close_document()
    return _f

###############################################################################


class Test_PyLOO_Document(unittest.TestCase):

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
        doc = pyloo.Document()
        file_name = os.getcwd() + "/test.ods"
        doc.open_document(file_name)
        field = doc.fields().field("TABLE_NAME")
        self.assertFalse(field.is_null(), "get field object")
        doc.close_document()

    @pyloo_open_close_doc
    def test_fields_count(self, doc):
        doc = pyloo.Document()
        file_name = os.getcwd() + "/test.ods"
        doc.open_document(file_name)
        self.assertEqual(doc.fields().count(), 7, "Wrong number of fields")
        doc.close_document()

###############################################################################


class Test_PyLOO_Field(unittest.TestCase):

    @pyloo_open_close_doc
    def test_field_set_get(self, doc):
        field = doc.fields().field("TABLE_NAME")
        test_value = "Test table name"
        self.assertTrue(field.set_value(test_value))
        self.assertEqual(field.value(), test_value)

    @pyloo_open_close_doc
    def test_field_insert_rows(self, doc):
        field = doc.fields().field("TABLE_NAME")
        test_value = "Test table name"
        self.assertTrue(field.set_value(test_value))
        self.assertEqual(field.value(), test_value)

###############################################################################


if __name__ == "__main__":
    unittest.main()
