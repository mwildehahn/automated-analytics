import setpath
import string

from numpy import *
from win32com.client import DispatchEx

import csv


class ExcelFatalError(Exception):
    """This is when excel totally blows up.

    I've only seen this happen a few times and it looks like its an issue with
    corrupted files or weird windows registery issues. I spent hours going
    through win32com library trying to dig into what was actually going on to
    no avail. Sometimes it works, other times it doesn't. When all else fails,
    write a custom exception class called 'FatalError'...

    """
    pass


class ExcelPathNotFound(Exception):
    def __init__(self, filename, action):
        self.filename = filename
        self.action = action

    def __str__(self):
        return 'Error %s file: %s, check to make sure path exists' % (
            self.action,
            self.filename,
        )


class BaseReferenceError(Exception):
    def __init__(self, reference):
        self.reference = reference


class InvalidSheet(BaseReferenceError):
    def __str__(self):
        return 'Error: referencing an invalid sheet, %s' % self.reference


class WorksheetNotFound(BaseReferenceError):
    def __str__(self):
        return 'Worksheet: %s, not found' % self.reference


class ExcelApi(object):
    def __init__(self, screen_updating=False, visible=False):
        self.excel = DispatchEx('Excel.Application')
        self.excel.ScreenUpdating = screen_updating
        self.excel.Visible = visible

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        self.quit(True)

    def quit(self, save=False):
        if hasattr(self, 'excel_book'):
            self.excel_book.Close(save)
            del self.excel_book
        if hasattr(self, 'excel'):
            self.excel.Quit()
            del self.excel

    def save(self):
        if hasattr(self, 'excel_book'):
            self.excel_book.Save()

    def save_as(self, filename):
        excel_type = self.get_exceltype(filename)
        if hasattr(self, 'excel_book'):
            try:
                self.excel_book.SaveAs(filename, excel_type)
            except:
                raise ExcelPathNotFound(filename, 'saving')

    def open_workbook(self, filename, add_book=True):
        try:
            self.excel_book = self.excel.Workbooks.Open(filename)
        except:
            if add_book:
                try:
                    self.excel_book = self.excel.Workbooks.Add()
                    self.save_as(filename)
                except:
                    self.quit()
                    raise ExcelPathNotFound(filename, 'writing')
            else:
                self.quit()
                raise ExcelPathNotFound(filename, 'opening')
        try:
            # hack to get around weird, weird, weird, WEIRD, occurrence where
            # some excel files that are opened will treat methods such as
            # 'Close' and 'Save' as properties, not methods. I guess it happens
            # often enough because they wrote the `_FlagAsMethod` method in the
            # first place...
            self.excel_book._FlagAsMethod('Close')
        except:
            raise ExcelFatalError((
                'Corrupted file. You must re-instantiate the ExcelApi '
                'with `visible`=True'
            ))

    def open_spreadsheet(self, sheet, filename=None, add_sheet=True):
        if not hasattr(self, 'excel_book'):
            if not filename:
                raise NameError('Filename must be specified')
            self.open_workbook(filename)

        try:
            sheet_num = int(sheet)
            try:
                self.excel_sheet = self.excel_book.Sheets(sheet_num)
            except:
                raise InvalidSheet(sheet_num)
        except ValueError:
            worksheet_names = [self.excel_book.Sheets(i).Name for i
                in range(1, self.excel_book.Sheets.Count + 1)]
            if sheet not in worksheet_names:
                if add_sheet:
                    print 'Worksheet, "%s", not found. Adding worksheet' % sheet
                    try:
                        self.excel_book.Sheets.Add(
                            After=self.excel_book.Sheets(
                                self.excel_book.Sheets.Count)).Name = sheet
                    except:
                        raise InvalidSheet(
                            '%s, is invalid. Add manually to see Excel Error' \
                            % sheet
                        )
                else:
                    raise WorksheetNotFound(sheet)
            self.excel_sheet = self.excel_book.Sheets(sheet)
        self.excel_sheet.Activate()

    def read(self, sheet, filename=None, quit=False):
        if not hasattr(self, 'excel_book'):
            if not filename:
                raise NameError('Filename must be provided')
            self.open_workbook(filename, add_book=add_book)

        self.open_spreadsheet(sheet)

        column_range = self.excel_sheet.Range('1:1')
        num_columns = int(self.excel.WorksheetFunction.CountA(column_range))
        headers = self.excel_sheet.Range(
            self.excel_sheet.Cells(1, 1),
            self.excel_sheet.Cells(1, num_columns),
        ).value

        workbook_dict = dict.fromkeys(headers[0])
        max_rows = 0
        for i in range(1, num_columns+1):
            num_rows = int(self.excel.WorksheetFunction.CountA(
                self.excel_sheet.Columns(i)))
            if num_rows > max_rows:
                max_rows = num_rows
        for i in range(1, num_columns+1):
            row_values = self.excel_sheet.Range(
                self.excel_sheet.Cells(1, i),
                self.excel_sheet.Cells(max_rows, i),
            ).value
            workbook_dict[row_values[0][0]] = [entry[0] for entry
                in row_values[1:]]
        if quit:
            self.quit()
        return workbook_dict

    def write(self, sheet, data, cellstr, filename=None, direction='h',
        quit=False, add_book=True):

        if not hasattr(self, 'excel_book'):
            if not filename:
                raise NameError('Filename must be provided')
            self.open_workbook(filename, add_book=add_book)

        self.open_spreadsheet(sheet)

        row, col = self.getcell(cellstr)
        output_dict, shp = self.data_export_cleaner(data, direction)
        a, b = shp
        start_cells = [(row, col+i) for i in range(b)]
        end_cells = [(row + a-1, col+i) for i in range(b)]
        for i in output_dict.keys():
            cell_range = eval('self.excel_sheet.Range(\
                self.excel_sheet.Cells%s, self.excel_sheet.Cells%s)' % \
                (start_cells[i], end_cells[i])
            )
            cell_range.Value = output_dict[i]
        if quit:
            self.quit(True)
        else:
            self.save()
        return

    def delete(self, sheet, cellrange, filename=None, quit=False):
        if not hasattr(self, 'excel_book'):
            if not filename:
                raise NameError('Filename must be provided')
            self.open_workbook(filename)

        self.open_spreadsheet(sheet, add_sheet=False)

        self.excel_sheet.Range(cellrange.upper()).Select()
        self.excel.selection.ClearContents()
        if quit:
            self.quit()

    def get_exceltype(self, filename):
        format_dict = {'xlsx':51, 'xlsm':52, 'xlsb':50, 'xls':56}
        character_dict = self.character_count(filename)
        if character_dict.get('.') > 1 or not character_dict.get('.'):
            raise NameError(
                'Error: Incorrect File Path Name, multiple or no periods')
        file_type = filename.split('.')[-1]
        if file_type not in format_dict.keys():
            raise NameError(
                'Error: Incorrect File Path, No excel file specified')
        else:
            return format_dict[file_type]

    def getcell(self, cell):
        '''Take a cell such as 'A1' and return the corresponding numerical row
        and column in excel'''
        cell_len = len(cell)
        temp_column = []
        temp_row = []
        row = []
        if cell_len < 2:
            raise NameError('Error, the cell you entered is not valid')
        for i in range(cell_len):
            if str.isdigit(cell[i]) == False:
                temp_column.append(cell[i])
            else:
                temp_row.append(cell[i])
        row.append(string.join(temp_row, ''))
        row = int(row[0])
        column = self.getnumericalcolumn(temp_column)
        return row, column

    @staticmethod
    def data_export_cleaner(data, direction):
        darray = array(data)
        shp = shape(darray)
        if len(shp) == 0:
            darray = array([data])
            darray = darray.reshape(1, 1)
        if len(shp) == 1:
            darray = array([data])
            if direction.lower() == 'v':
                darray = darray.transpose()
        shp = shape(darray)
        output_dict = {}
        for i in range(shp[1]):
            output_dict[i] = [(str(darray[j, i]),) for j in range(shp[0])]
        return output_dict, shp

    @staticmethod
    def character_count(a_string):
        character_dict = dict()
        for c in a_string:
            character_dict[c] = character_dict.get(c, 0) + 1
        return character_dict

    @staticmethod
    def getnumericalcolumn(column):
        '''Take an excel column specification such as 'A' and return its
        numerical equivalent in excel'''
        alpha = str(string.ascii_uppercase)
        alphadict = dict(zip(alpha, range(1, len(alpha) + 1)))
        if len(column) == 1:
            numcol = alphadict[column[0]]
        elif len(column) == 2:
            numcol = alphadict[column[0]] * 26 + alphadict[column[1]]
        elif len(column) == 3:
            numcol = 26 ** 2 + alphadict[column[1]] * 26 + alphadict[column[2]]
        return numcol

def xlsxwrite(filename, sheet, data, cellstr, direction='h'):
    """
    Bundle write method for backwards compatability with existing scripts
    """
    with ExcelApi() as excel_api:
        excel_api.open_workbook(filename)
        excel_api.write(sheet, data, cellstr, direction=direction)

def xlsxdelete(filename, sheet, cellrange):
    """
    Bundle delete method for backwards compatability with existing scripts
    """
    with ExcelApi() as excel_api:
        excel_api.open_workbook(filename)
        excel_api.delete(sheet, cellrange)

def xlsxread(filename, sheet):
    """
    Bundle write method for backwards compatability with existing scripts
    """
    with ExcelApi() as excel_api:
        excel_api.open_workbook(filename)
        result = excel_api.read(sheet)
    return result

def readcsv(filename, headers=True):
    """Function for reading a csv file.

    Returns a dictionary with the file headers as dictionary keys.

    """
    reader = csv.reader(open(filename, 'rU'), dialect = 'excel')
    contents = []
    for line in reader:
        contents.append(line)
    labels = contents[0]
    dictionary = dict.fromkeys(labels)
    for key in dictionary:
        dictionary[key] = []
    for i in range(1,len(contents)):
        for j in range(len(labels)):
            dictionary[labels[j]].append(contents[i][j])
    return dictionary
