#!/usr/bin/python
#
# Copyright (C) 2012 Yoav Aviram.
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#      http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
import gdata.spreadsheet.service


ID_FIELD = '__rowid__'


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
        self.query = None

    def _row_to_dict(self, row):
        """Turn a row of values into a dictionary.
        :param row:
            A row element.
        :return:
            A dictionary with rows.
        """
        result = dict([(key, row.custom[key].text) for key in row.custom])
        result[ID_FIELD] = row.id.text.split('/')[-1]
        return result

    def _get_row_entries(self, query=None):
        """Get Row Entries.

        :return:
            A rows entry.
        """
        if not self.entries:
            self.entries = self.gd_client.GetListFeed(
                query=query, **self.keys).entry
        return self.entries

    def _get_row_entry_by_id(self, id):
        """Get Row Entry by ID

        First search in cache, then fetch.
        :param id:
            A string row ID.
        :return:
            A row entry.
        """
        entry = [entry for entry in self._get_row_entries()
                 if entry.id.text.split('/')[-1] == id]
        if not entry:
            entry = self.gd_client.GetListFeed(row_id=id, **self.keys).entry
            if not entry:
                raise WorksheetException("Row ID '{0}' not found.").format(id)
        return entry[0]

    def _flush_cache(self):
        """Flush Entries Cache."""
        self.entries = None

    def _make_query(self, query=None, order_by=None, reverse=None):
        """Make Query.

         A utility method to construct a query.

        :return:
            A :class:`~,gdata.spreadsheet.service.ListQuery` or None.
        """
        if query or order_by or reverse:
            q = gdata.spreadsheet.service.ListQuery()
            if query:
                q.sq = query
            if order_by:
                q.orderby = order_by
            if reverse:
                q.reverse = reverse
            return q
        else:
            return None

    def get_rows(self, query=None, order_by=None,
                 reverse=None, filter_func=None):
        """Get Rows

        :param query:
            A string structured query on the full text in the worksheet.
              [columnName][binaryOperator][value]
              Supported binaryOperators are:
              - (), for overriding order of operations
              - = or ==, for strict equality
              - <> or !=, for strict inequality
              - and or &&, for boolean and
              - or or ||, for boolean or.
        :param order_by:
            A string which specifies what column to use in ordering the
            entries in the feed. By position (the default): 'position' returns
            rows in the order in which they appear in the GUI. Row 1, then
            row 2, then row 3, and so on. By column:
            'column:columnName' sorts rows in ascending order based on the
            values in the column with the given columnName, where
            columnName is the value in the header row for that column.
        :param reverse:
            A string which specifies whether to sort in descending or ascending
            order.Reverses default sort order: 'true' results in a descending
            sort; 'false' (the default) results in an ascending sort.
        :param filter_func:
            A lambda function which applied to each row, Gets a row dict as
            argument and returns True or False. Used for filtering rows in
            memory (as opposed to query which filters on the service side).
        :return:
            A list of row dictionaries.
        """
        new_query = self._make_query(query, order_by, reverse)
        if self.query is not None and self.query != new_query:
            self._flush_cache()
        self.query = new_query
        rows = [self._row_to_dict(row)
            for row in self._get_row_entries(query=self.query)]
        if filter_func:
            rows = filter(filter_func, rows)
        return rows

    def update_row(self, row_data):
        """Update Row (By ID).

        Only the fields supplied will be updated.
        :param row_data:
            A dictionary containing row data. The row will be updated according
            to the value in the ID_FIELD.
        :return:
            The updated row.
        """
        try:
            id = row_data[ID_FIELD]
        except KeyError:
            raise WorksheetException("Row does not contain '{0}' field. "
                                "Please update by index.".format(ID_FIELD))
        entry = self._get_row_entry_by_id(id)
        new_row = self._row_to_dict(entry)
        new_row.update(row_data)
        entry = self.gd_client.UpdateRow(entry, new_row)
        if not isinstance(entry, gdata.spreadsheet.SpreadsheetsList):
            raise WorksheetException("Row update failed: '{0}'".format(entry))
        for i, e in enumerate(self.entries):
            if e.id.text == entry.id.text:
                self.entries[i] = entry
        return self._row_to_dict(entry)

    def update_row_by_index(self, index, row_data):
        """Update Row By Index

        :param index:
            An integer designating the index of a row to update (zero based).
            Index is relative to the returned result set, not to the original
            spreadseet.
        :param row_data:
            A dictionary containing row data.
        :return:
            The updated row.
        """
        entry = self._get_row_entries(self.query)[index]
        row = self._row_to_dict(entry)
        row.update(row_data)
        entry = self.gd_client.UpdateRow(entry, row)
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

    def delete_row(self, row):
        """Delete Row (By ID).

        Requires that the given row dictionary contains an ID_FIELD.
        :param row:
            A row dictionary to delete.
        """
        try:
            id = row[ID_FIELD]
        except KeyError:
            raise WorksheetException("Row does not contain '{0}' field. "
                                "Please delete by index.".format(ID_FIELD))
        entry = self._get_row_entry_by_id(id)
        self.gd_client.DeleteRow(entry)
        for i, e in enumerate(self.entries):
            if e.id.text == entry.id.text:
                del self.entries[i]

    def delete_row_by_index(self, index):
        """Delete Row By Index

        :param index:
            A row index. Index is relative to the returned result set, not to
            the original spreadsheet.
        """
        entry = self._get_row_entries(self.query)[index]
        self.gd_client.DeleteRow(entry)
        del self.entries[index]

    def delete_all_rows(self):
        """Delete All Rows
        """
        entries = self._get_row_entries(self.query)
        for entry in entries:
            self.gd_client.DeleteRow(entry)
        self._flush_cache()
