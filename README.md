Google Spreadsheet API
========================
A simple Python wrapper for the Google Spreadsheet API.


Features
--------

* An object oriented interface to Worksheets.
* Supports List Feed view of spreadsheet rows, represented as dictionaries.
* Compatible with Google App Engine


Requirements
--------------
Before you get started, make sure you have:

* Installed [Gdata](http://code.google.com/p/gdata-python-client/) (pip install gdata)

Usage
-----

List Spreadsheets and Worksheets:

     >>> from google_spreadsheet.api import SpreadsheetAPI
     >>> api = SpreadsheetAPI(GOOGLE_SPREADSHEET_USER,
                     GOOGLE_SPREADSHEET_PASSWORD, GOOGLE_SPREADSHEET_SOURCE)
     >>> spreadsheets = api.list_spreadsheets()
     >>> spreadsheets
     [('MyFirstSpreadsheet', 'tkZQWzwHEjKTWFFCAgw'), ('MySecondSpreadsheet', 't5I-ZPGdXjTrjMefHcg'), ('MyThirdSpreadsheet', 't0heCWhzCmm9Y-GTTM_Q')]
     >>> worksheets = spreadsheet.list_worksheets(spreadsheets[0][1])
     >>> worksheets
     [('MyFirstWorksheet', 'od7'), ('MySecondWorksheet', 'od6'), ('MyThirdWorksheet', 'od4')]

Please note that in order to work with a Google Spreadsheet it must be accessible
to the user who's login credentials are provided. The `GOOGLE_SPREADSHEET_SOURCE`
is used by google to identify your application and track API calls.

Working with a Worksheet:

    >>> sheet = spreadsheet.get_worksheet('tkZQWzwHEjKTWFFCAgw', 'od7')
    >>> rows = sheet.get_rows()
    >>> len(rows)
    18
    >>> row_to_update = rows[0]
    >>> row_to_update['name'] = 'New Name'
    >>> sheet.update_row(0, row_to_update)
    {'name': 'New Name'...}
    >>> row_to_insert = rows[0]
    >>> row_to_insert['name'] = 'Another Name'
    >>> sheet.insert_row(row_to_insert)
    {'name': 'Another Name'...}
    >>> sheet.delete_row(18)
    >>> sheet.delete_all_rows()

That's it.

For more information about these calls, please consult the [Google Spreadsheets
API Developer Guide](https://developers.google.com/google-apps/spreadsheets/).


License
-------

Copyright &copy; 2012 Yoav Aviram

See LICENSE for details.

