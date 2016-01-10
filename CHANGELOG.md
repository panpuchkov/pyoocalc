# Change Log
All notable changes to this project will be documented in this file.

## [0.0.2] - 2016-01-09
### Added
``autostart`` option in the constructor.

Auto starts Libre/Open Office with a listening socket.

Example:

doc = pyloo.Document(autostart=True)

## [0.0.3] - 2016-01-10
### Changed
Some class  methods are changed to be as properties via decorator
``@property``.

Old code:

doc.sheets().insert_spreadsheet("Test1", 0)

New code:

doc.sheets.insert_spreadsheet("Test1", 0)

Changed methods:

* Document.o_doc() -> Document.o_doc

* Document.sheets() -> Document.sheets

* Document.fields() -> Document.fields

* Fields.document() -> Fields.document

* Sheets.o_sheets() -> Sheets.o_sheets
