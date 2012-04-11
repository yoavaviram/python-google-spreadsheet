import gdata.spreadsheet.service


class WorksheetException(Exception):
    """Base class for spreadsheet exceptions.
    """
    pass


class SpreadsheetAPI(object):
    def __init__(self, email, password, source):
        """Initialise a Spreadsheet API wrapper.

        :param email:
            A string representing a google login email.
        :param password:
            A string representing a google login password.
        :param source:
            A string representing source (much like a user agent).
        """
        self.email = email
        self.password = password
        self.source = source

    def _get_client(self):
        """Initialize a `gdata` client.

        :returns:
            A gdata client.
        """
        gd_client = gdata.spreadsheet.service.SpreadsheetsService()
        gd_client.email = self.email
        gd_client.password = self.password
        gd_client.source = self.source
        gd_client.ProgrammaticLogin()
        return gd_client

    def list_spreadsheets(self):
        """List Spreadsheets.

        :return:
            A list with information about the spreadsheets available
        """
        sheets = self._get_client().GetSpreadsheetsFeed()
        return map(lambda e: (e.title.text, e.id.text.rsplit('/', 1)[1]),
            sheets.entry)

    def list_worksheets(self, spreadsheet_key):
        """List Spreadsheets.

        :return:
            A list with information about the spreadsheets available
        """
        wks = self._get_client().GetWorksheetsFeed(
            key=spreadsheet_key)
        return map(lambda e: (e.title.text, e.id.text.rsplit('/', 1)[1]),
            wks.entry)

    def get_worksheet(self, spreadsheet_key, worksheet_key):
        """Get Worksheet.

        :param spreadsheet_key:
            A string representing a google spreadsheet key.
        :param worksheet_key:
            A string representing a google worksheet key.
        """
        return Worksheet(self._get_client(), spreadsheet_key, worksheet_key)


class Worksheet(object):
    """Worksheet wrapper class.
    """
    def __init__(self, gd_client, spreadsheet_key, worksheet_key):
        """Initialise a client

        :param gd_client:
            A GDATA client.
        :param spreadsheet_key:
            A string representing a google spreadsheet key.
        :param worksheet_key:
            A string representing a google worksheet key.
        """
        self.gd_client = gd_client
        self.spreadsheet_key = spreadsheet_key
        self.worksheet_key = worksheet_key
        self.keys = {'key': spreadsheet_key, 'wksht_id': worksheet_key}
        self.entries = None

    def flush_entry_cache(self):
        self.entries = None

    def _row_to_dict(self, row):
        """Turn a row of values into a dictionary.
        :param row:
            A row element.
        :return:
            A dict.
        """
        return dict([(key, row.custom[key].text) for key in row.custom])

    def _get_row_entries(self):
        """Get Row Entries

        :return:
            A rows entry
        """
        if not self.entries:
            self.entries = self.gd_client.GetListFeed(**self.keys).entry
        return self.entries

    def get_rows(self):
        """Get Rows

        :return:
            A list of rows dictionaries.
        """
        return [self._row_to_dict(row) for row in self._get_row_entries()]

    def update_row(self, index, row_data):
        """Update Row

        :param index:
            An integer designating the index of a row to update.
        :param row_data:
            A dictionary containing row data.
        :return:
            A row dictionary for the updated row.
        """
        entries = self._get_row_entries()
        rows = self.get_rows()
        rows[index].update(row_data)
        entry = self.gd_client.UpdateRow(entries[index], rows[index])
        if not isinstance(entry, gdata.spreadsheet.SpreadsheetsList):
            raise WorksheetException("Row update failed: '{0}'".format(entry))
        self.entries[index] = entry
        return self._row_to_dict(entry)

    def insert_row(self, row_data):
        """Insert Row

        :param row_data:
            A dictionary containing row data.
        :return:
            A row dictionary for the inserted row.
        """
        entry = self.gd_client.InsertRow(row_data, **self.keys)
        if not isinstance(entry, gdata.spreadsheet.SpreadsheetsList):
            raise WorksheetException("Row insert failed: '{0}'".format(entry))
        if self.entries:
            self.entries.append(entry)
        return self._row_to_dict(entry)

    def delete_row(self, index):
        """Delete Row

        :param index:
            A row index.
        """
        entries = self._get_row_entries()
        self.gd_client.DeleteRow(entries[index])
        if self.entries:
            del self.entries[index]

    def delete_all_rows(self):
        """Delete All Rows
        """
        entries = self._get_row_entries()
        for entry in entries:
            self.delete_row(entry)
