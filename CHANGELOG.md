# Change Log
All notable changes to this project will be documented in this file.

## [0.0.4] - 2016-01-11
### Changed
- Some class  methods are changed into a properties via the decorator 
``@property``:
    - Document.is_null() -> Document.is_null
    - Document.version() -> Document.version
    - Fields.is_null() -> Fields.is_null
    - Fields.count() -> Fields.count
    - Field.is_null() -> Field.is_null
    - Sheets.is_null() -> Sheets.is_null
    - Sheets.count() -> Sheets.count
    - Sheet.is_null() -> Sheet.is_null
- In the function ``start_office_instance`` print output changed into 
exception raise on errors.
- Check function arguments and raise a ``ValueError`` on errors.
- Add ``__enter__`` and ``__exit__`` methods to the ``Document`` class 
according to [PEP 0343](https://www.python.org/dev/peps/pep-0343/) (the “with” 
statement).

Sorry for no back compatibility, I guess that's was the last time.

## [0.0.3] - 2016-01-10
### Changed
Some class  methods are changed to be as properties via decorator
``@property``.

#####Old code:
doc.sheets().insert_spreadsheet("Test1", 0)

doc.fields().field("FIELD_1")

#####New code:
doc.sheets.insert_spreadsheet("Test1", 0)

doc.fields.field("FIELD_1")

####Changed methods:
- Document.o_doc() -> Document.o_doc
- Document.sheets() -> Document.sheets
- Document.fields() -> Document.fields
- Fields.document() -> Fields.document
- Sheets.o_sheets() -> Sheets.o_sheets


## [0.0.2] - 2016-01-09
### Added
``autostart`` option in the constructor.

Auto starts Libre/Open Office with a listening socket.

Example:

doc = pyoocalc.Document(autostart=True)
