# coding: utf-8

import unittest
###############################################################################
import os
import sys
sys.path.append('./../')

###############################################################################
import pyloo

###############################################################################


class Test_PyLOO_Document(unittest.TestCase):

    def test_document_initialize(self):
        doc = pyloo.Document()
        self.assertFalse(doc.is_null())

    def test_document_open_close(self):
        doc = pyloo.Document()
        file_name = os.getcwd() + "/test.ods"
        self.assertTrue(doc.open_document(file_name), "OPEN")
        self.assertTrue(doc.close_document(), "CLOSE")

    def test_document_new_close(self):
        doc = pyloo.Document()
        self.assertTrue(doc.new_document(), "NEW")
        self.assertTrue(doc.close_document(), "CLOSE")

    def test_document_new_save_close(self):
        doc = pyloo.Document()
        file_name_saved = os.getcwd() + "/test_saved.ods"
        self.assertTrue(doc.new_document(), "OPEN")
        self.assertTrue(doc.save_document(file_name_saved), "SAVE")
        self.assertTrue(os.path.isfile(file_name_saved), "Is file exists")
        os.remove(file_name_saved)
        self.assertTrue(doc.close_document(), "CLOSE")

    def test_document_sheets(self):
        doc = pyloo.Document()
        file_name = os.getcwd() + "/test.ods"
        doc.open_document(file_name)
        sheets = doc.sheets()
        self.assertFalse(sheets.is_null(), "get sheets object")
        doc.close_document()

    def test_document_fields(self):
        doc = pyloo.Document()
        file_name = os.getcwd() + "/test.ods"
        doc.open_document(file_name)
        fields = doc.fields()
        self.assertFalse(fields.is_null(), "get fields object")
        doc.close_document()

###############################################################################


class Test_PyLOO_Fields(unittest.TestCase):

    def test_fields_field(self):
        doc = pyloo.Document()
        file_name = os.getcwd() + "/test.ods"
        doc.open_document(file_name)
        field = doc.fields().field("TABLE_NAME")
        self.assertFalse(field.is_null(), "get field object")
        doc.close_document()

    def test_fields_count(self):
        doc = pyloo.Document()
        file_name = os.getcwd() + "/test.ods"
        doc.open_document(file_name)
        self.assertEqual(doc.fields().count(), 7, "Wrong number of fields")
        doc.close_document()

###############################################################################


class Test_PyLOO_Field(unittest.TestCase):

    def test_field_set_get(self):
        doc = pyloo.Document()
        file_name = os.getcwd() + "/test.ods"
        doc.open_document(file_name)
        field = doc.fields().field("TABLE_NAME")

        test_value = "Test table name"
        self.assertTrue(field.set_value(test_value))
        self.assertEqual(field.value(), test_value)

        doc.close_document()

    def test_field_insert_rows(self):
        doc = pyloo.Document()
        file_name = os.getcwd() + "/test.ods"
        doc.open_document(file_name)
        field = doc.fields().field("TABLE_NAME")

        test_value = "Test table name"
        self.assertTrue(field.set_value(test_value))
        self.assertEqual(field.value(), test_value)

        doc.close_document()

###############################################################################


if __name__ == "__main__":
    unittest.main()
