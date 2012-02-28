import gdata.spreadsheet.service as gdata_spreadsheet
from robo_configs import (
    robo_user,
    robo_passwd,
)

#TODO

#As of right now, this api only reads data. If I get time it would be cool to
#make this write as well so reports could live soley on gdocs.


class GoogleDocsApi(object):
    """ API for interacting with google docs.

    Right now only supports reading given a specific document title.
    The email defined in the __init__ must have access to the document,
    defaults to robo.analyst

    """

    def __init__(self, document_title=None, email=robo_user,
        password=robo_passwd, source='robo_analyst', worksheet_num=1):
        self._email = email
        self._password = password
        self._source = source
        self._document_title = document_title
        self._current_ids = {'spreadsheet_id': None, 'worksheet_id': None}
        self._rows = None

        # Activate the google docs api
        self.read_client = gdata_spreadsheet.SpreadsheetsService()
        self.read_client.email = self._email
        self.read_client.password = self._password
        self.read_client.source = self._source
        self.read_client.ProgrammaticLogin()

        if self._document_title:
            self._rows = self.get_worksheet_rows(self._document_title,
                worksheet_num=worksheet_num)

    def get_spreadsheet_feed(self, document_title, exact=True):
        doc_query = gdata_spreadsheet.DocumentQuery()
        doc_query['title'] = document_title
        if exact:
            doc_query['title-exact'] = 'true'
        spreadsheet_feed = self.read_client.GetSpreadsheetsFeed(query=doc_query)
        spreadsheet_id = spreadsheet_feed.entry[0].id.text.rsplit('/', 1)[1]
        return spreadsheet_feed, spreadsheet_id

    def get_worksheet_feed(self, spreadsheet_id=None, document_title=None,
        exact=True, worksheet_num=1):
        if not any([spreadsheet_id, document_title]):
            raise NameError('Must provide either worksheet_id \
                or document_title')
        if not spreadsheet_id:
            spreadsheet_feed, spreadsheet_id = self.get_spreadsheet_feed(
                document_title)
        worksheet_feed = self.read_client.GetWorksheetsFeed(spreadsheet_id)
        worksheet_id = worksheet_feed.entry[worksheet_num - 1] \
            .id.text.rsplit('/', 1)[1]
        return worksheet_feed, worksheet_id

    def get_worksheet_rows(self, document_title=None, spreadsheet_id=None,
        worksheet_id=None, worksheet_num=1):
        if not all([spreadsheet_id, worksheet_id]) and not document_title:
            raise NameError('Must provide either spreadsheet & worksheet \
                ids or document_title')
        if not all([spreadsheet_id, worksheet_id]):
            spreadsheet_feed, spreadsheet_id = self.get_spreadsheet_feed(
                document_title)
            worksheet_feed, worksheet_id = self.get_worksheet_feed(
                spreadsheet_id, worksheet_num=worksheet_num)

        # To allow for easy updating of feed & dict output
        self._current_ids['spreadsheet_id'] = spreadsheet_id
        self._current_ids['worksheet_id'] = worksheet_id
        self._rows = self.read_client.GetListFeed(spreadsheet_id,
            worksheet_id).entry

        return self._rows

    def get_updated_rows(self):
        if not all(self._current_ids.values()):
            raise NameError('No feed is active')
        updated_rows = self.get_worksheet_rows(
            spreadsheet_id=self._current_ids['spreadsheet_id'],
            worksheet_id=self._current_ids['worksheet_id'],
        )
        self._rows = updated_rows
        return updated_rows

    @property
    def dict_output(self):
        if not self._rows:
            return {}
        data_dict = dict((key, []) for key in self._rows[0].custom)
        for row in self._rows:
            for key in row.custom:
                data_dict[key].append(row.custom[key].text)
        return data_dict
