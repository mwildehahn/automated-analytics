import gdata.spreadsheet
import gdata.spreadsheet.service as gdata_spreadsheet
import setpath
from numpy import column_stack
from logicals import (
    stack_dict,
    sort_table,
)
from collections import defaultdict
from django_utils import SortedDict
from robo_configs import (
    robo_user,
    robo_passwd,
)

# Error when trying to use InsertRow
# gdata.service.RequestError: {'status': 400, 'body': 'We&#39;re sorry, a
# server error occurred. Please wait a bit and try reloading your
# spreadsheet.','reason': 'Bad Request'}

# means we have to use exisitng cells

class GoogleDocsApi(object):
    """ API for interacting with google docs.

    Right now only supports reading given a specific document title.
    The email defined in the __init__ must have access to the document,
    defaults to robo.analyst.

    """

    def __init__(
        self,
        document_title=None,
        email=robo_user,
        password=robo_passwd,
        source='robo_analyst',
        worksheet_num=1,
    ):
        self._email = email
        self._password = password
        self._source = source
        self._document_title = document_title
        self._rows = None

        # Activate the google docs api
        self.g_client = gdata_spreadsheet.SpreadsheetsService()
        self.g_client.email = self._email
        self.g_client.password = self._password
        self.g_client.source = self._source
        self.g_client.ProgrammaticLogin()

        if self._document_title:
            self._rows = self.get_worksheet_rows(self._document_title,
                worksheet_num=worksheet_num)

    def get_spreadsheet_feed(self, document_title, exact=True, force=False):

        # load from cache unless we force
        if hasattr(self, '_spreadsheet_feed') and not force:
            return self._spreadsheet_feed, self._spreadsheet_id

        doc_query = gdata_spreadsheet.DocumentQuery()
        doc_query['title'] = document_title
        if exact:
            doc_query['title-exact'] = 'true'
        spreadsheet_feed = self.g_client.GetSpreadsheetsFeed(
            query=doc_query,
        )
        if not spreadsheet_feed:
            raise ValueError('%s does not exist' % document_title)
        spreadsheet_id = spreadsheet_feed.entry[0].id.text.rsplit('/', 1)[1]

        # cache spreadsheet_feed and spreadsheet_id
        self._spreadsheet_feed = spreadsheet_feed
        self._spreadsheet_id = spreadsheet_id

        return spreadsheet_feed, spreadsheet_id

    def get_worksheet_feed(
        self,
        spreadsheet_id=None,
        document_title=None,
        exact=True,
        worksheet_num=1,
        force=False,
    ):

        # load from cache unless we force
        if hasattr(self, '_worksheet_feed') and not force:
            return self._worksheet_feed, self._worksheet_id

        if not hasattr(self, '_spreadsheet_id') and not spreadsheet_id:
            self.get_spreadsheet_feed(
                document_title,
                exact,
            )

        worksheet_feed = self.g_client.GetWorksheetsFeed(self._spreadsheet_id)
        worksheet_id = (
            worksheet_feed.entry[worksheet_num - 1].id.text.rsplit('/', 1)[1]
        )

        # cache the worksheet_feed and worksheet_id
        self._worksheet_feed = worksheet_feed
        self._worksheet_id = worksheet_id

        return worksheet_feed, worksheet_id

    def get_cells_feed(
        self,
        spreadsheet_id=None,
        document_title=None,
        exact=True,
        worksheet_num=1,
        worksheet_id=None,
        force=False,
    ):
        # load from cache unless we force
        if hasattr(self, '_cells_feed') and not force:
            return self._cells_feed

        if (
            not hasattr(self, '_spreadsheet_id')
            or not hasattr(self, '_worksheet_id')
        ) or (
            not spreadsheet_id or not worksheet_id
        ):
            self.get_worksheet_feed(
                spreadsheet_id,
                document_title,
                exact,
                worksheet_num,
                force=force,
            )

        cells_feed = self.g_client.GetCellsFeed(
            self._spreadsheet_id,
            self._worksheet_id,
        )

        # cache the cells_feed
        self._cells_feed = cells_feed
        return cells_feed

    def get_worksheet_rows(
        self,
        document_title=None,
        spreadsheet_id=None,
        worksheet_id=None,
        worksheet_num=1,
        exact=True,
        force=False,
    ):
        if (
            not hasattr(self, '_spreadsheet_id')
            or not hasattr(self, '_worksheet_id')
        ) or (
            not spreadsheet_id or not worksheet_id
        ):
            self.get_worksheet_feed(
                spreadsheet_id=spreadsheet_id,
                document_title=document_title,
                exact=exact,
                worksheet_num=worksheet_num,
                force=force,
            )

        # cache the rows
        self._rows = self.g_client.GetListFeed(
            self._spreadsheet_id,
            self._worksheet_id,
        ).entry

        return self._rows

    def get_updated_rows(self):

        if not all(self._current_ids.values()):
            raise NameError('No feed is active')

        updated_rows = self.get_worksheet_rows(
            spreadsheet_id=self._spreadsheet_id,
            worksheet_id=self._worksheet_id,
        )
        self._rows = updated_rows
        return updated_rows

    def update_spreadsheet(
        self,
        data_dict,
        headers=None,
        spreadsheet_id=None,
        worksheet_id=None,
        document_title=None,
        worksheet_num=1,
        exact=True,
        cells_feed=None,
        force=False,
    ):
        if (
            not hasattr(self, '_spreadsheet_id')
            or not hasattr(self, '_worksheet_id')
        ) or (
            not spreadsheet_id or not worksheet_id
        ) and not cells_feed:
            self.get_cells_feed(
                spreadsheet_id=spreadsheet_id,
                document_title=document_title,
                exact=exact,
                worksheet_num=worksheet_num,
                force=force,
            )
        if cells_feed:
            self._cells_feed = cells_feed

        if not headers:
            # if headers aren't given, the headers won't be sorted
            headers = data_dict.keys()

        # store column referece
        column = 1
        for header in headers:
            print 'processing %s...' % header

            # use a batch request to limit the number of requests
            batch_request = gdata.spreadsheet.SpreadsheetsCellsFeed()

            try:
                column_data = data_dict[header]
            except KeyError:
                raise KeyError('Provided headers must match in the data_dict')
            else:
                cell_map = self.get_column_map()
                try:
                    column_map_data = cell_map[column]
                except KeyError:
                    raise KeyError(
                        'Column mapping not available for column: %d' % column
                    )
                else:
                    column_map_data[0].cell.inputValue = header
                    batch_request.AddUpdate(column_map_data[0])
                    for index, field in enumerate(column_data):
                        try:
                            # adding 1 for column_map_data index because we
                            # wrote the header already
                            column_map_data[index + 1].cell.inputValue = (
                                column_data[index]
                            )
                        except IndexError:
                            raise IndexError(
                                'The sheet you\'re updating doesn\'t have '
                                'enough availalbe cells'
                            )
                        batch_request.AddUpdate(column_map_data[index + 1])
                    # execute the batch request for this column
                    self.g_client.ExecuteBatch(
                        batch_request,
                        self._cells_feed.GetBatchLink().href,
                    )
            column += 1

    def safe_clear_cells(self, cells_feed=None):
        if not hasattr(self, '_cells_feed') and not cells_feed:
            raise NameError('No cached `_cells_feed`, provide one')
        if cells_feed:
            self._cells_feed = cells_feed

        cell_map = self.get_column_map()
        for values in cell_map.values():
            # ignoring columns without headers
            if values[0].cell.text == '-':
                print '! skipping column without header'
                continue
            batch_request = gdata.spreadsheet.SpreadsheetsCellsFeed()
            for value in values[1:]:
                # make value '-' so that we can safely clear the file but still
                # update with values later
                value.cell.inputValue = '-'
                batch_request.AddUpdate(value)
            # execute the batch request for this column
            self.g_client.ExecuteBatch(
                batch_request,
                self._cells_feed.GetBatchLink().href,
            )


    def get_column_map(self, cells_feed=None):
        if not hasattr(self, '_cells_feed') and not cells_feed:
            raise NameError('No cached `_cells_feed`, provide one')
        if cells_feed:
            self._cells_feed = cells_feed
        column_map = defaultdict(list)
        for field in self._cells_feed.entry:
            column_map[int(field.cell.col)].append(field)
        return column_map

    def get_row_map(self, cells_feed=None):
        if not hasattr(self, '_cells_feed') and not cells_feed:
            raise NameError('No cached `_cells_feed`, provide one')
        if cells_feed:
            self._cells_feed = cells_feed
        row_map = defaultdict(list)
        for field in self._cells_feed.entry:
            row_map[int(field.cell.row)].append(field)
        return row_map

    def get_headers(self, cells_feed=None):
        if not hasattr(self, '_cells_feed') and not cells_feed:
            raise NameError('No cached `_cells_feed`, provide one')
        if cells_feed:
            self._cells_feed = cells_feed
        row_map = self.get_row_map()
        header_row = row_map[1]
        return [field.cell.text for field in header_row]

    def get_converted_row_dict(self, cells_feed=None):
        if not hasattr(self, '_cells_feed') and not cells_feed:
            raise NameError('No cached `_cells_feed`, provide one')
        if not self._rows:
            return {}
        if cells_feed:
            self._cells_feed = cells_feed
        headers = self.get_headers()
        data_dict = SortedDict((key, '') for key in headers)
        temp_dict = self.rows_dict_output
        # gdata strips out some characters from their headers and doesn't have
        # them sorted properly
        for key in data_dict:
            data_dict[key] = temp_dict[self._convert_to_gdata_header(key)]
        return data_dict

    @property
    def rows_dict_output(self):
        if not self._rows:
            return {}
        data_dict = defaultdict(list)
        for row in self._rows:
            for key in row.custom:
                # can't write `None` to spreadsheet because we wouldn't be able
                # to update those cells
                if row.custom[key].text is None:
                    field_value = '-'
                else:
                    field_value = row.custom[key].text
                data_dict[key].append(field_value)
        return data_dict

    @property
    def cells_dict_output(self):
        if not hasattr(self, '_cells_feed'):
            return {}
        headers = self.get_headers()
        data_dict = SortedDict((key, []) for key in headers)
        column_map = self.get_column_map()
        # match up columns to headers
        for column, column_data in column_map.iteritems():
            if column_data:
                data_dict[column_data[0].cell.text] = (
                    [field.cell.text for field in column_data[1:]]
                )
        return data_dict

    def _convert_to_gdata_header(self, header):
        return header.replace('/', '').replace(' ', '').lower()
