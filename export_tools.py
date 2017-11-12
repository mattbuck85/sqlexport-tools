import xlsxwriter
import unicodecsv as csv
from datetime import datetime,date


class DatabaseExportEmptyRowsException(Exception):
    pass

class DatabaseExportHelperDuplicates(Exception):
    pass

class DatabaseExport(object):
    '''Controller class for the document writer tools below.
       Pass it any python database cursor and get a column dump.
       Tested with MSSQL and MySQL.'''

    def __init__(self, cursor, table, custom_sql=None):
        self.cursor = cursor
        self.table = table
        self.rows = self.get_rows(custom_sql)
        self.column_list = self.get_column_list()

    def get_rows(self, custom_sql=None):
        if custom_sql:
            sql = custom_sql
        else:
            sql = 'SELECT * FROM %s' % self.table
        self.cursor.execute(sql)
        rows = self.cursor.fetchall()
        if len(rows) == 0:
            raise DatabaseExportEmptyRowsException('No rows returned from query. SQL:%s' % sql)
        return rows

    def get_column_list(self):
        column_list = []
        for column_description in self.cursor.description:
            column_list.append(column_description[0])
        return column_list

    '''Instantiate a writer class and pass it to export'''
    def export(self,writer,**kwargs):
        writer.perform(self.column_list, self.rows, **kwargs)

class CsvWriterTool(object):
    '''When passed to a DatabaseExport.export, perform a dump to CSV'''

    def __init__(self, file_handle, dialect='excel', **kwargs):
        self.file_handle = file_handle
        self.dialect = dialect

    def get_row_dicts(self, column_list, rows):
        row_dicts = []
        for row in rows:
            row_dicts.append(dict(zip(column_list, row)))
        return row_dicts

    def perform(self, column_list, rows, **kwargs):
        self.export(column_list, rows)

    def export(self, column_list, rows):
        self.row_dicts = self.get_row_dicts(column_list, rows)
        self.csvfile = self.file_handle
        table_dict_writer = csv.DictWriter(self.csvfile, fieldnames=column_list, dialect=self.dialect)
        table_dict_writer.writeheader()
        for row_dict in self.row_dicts:
            table_dict_writer.writerow(row_dict)

    def close(self):
        self.csvfile.close()

'''Default parameters for XlsxWriter.  Unpack into the constructor.'''
xlsx_default_kwargs = {
                        'string_encoding':'windows-1252',
                        'default_date_format':'YYYY-MM-DD',
                        'column_format':{'bold':True},
                      }

class XlsxWriterTool(object):
    '''When passed to a DatabseExport.export, perform a dump to xlsx
       Will make any null columns blank.  Experimental vlookup feature for
       multi-worksheet dumps.  Workbook kwargs: http://xlsxwriter.readthedocs.io/workbook.html'''

    def __init__(self,filename, string_encoding=None, blank_nulls=True, column_format=None, **workbook_kwargs):
        self.string_encoding = string_encoding
        self.blank_nulls = blank_nulls
        self.filename = filename
        self.workbook = xlsxwriter.Workbook(self.filename,workbook_kwargs)
        self.date_format = workbook_kwargs.get('default_date_format',None)
        self.column_format = self.workbook.add_format(column_format)
        self.worksheets = {}

    def get_num_format(self, format_string):
        if format_string:
            num_format = self.workbook.add_format()
            num_format.set_num_format(format_string)
        else:
            num_format = None
        return num_format

    def create_worksheet(self, name):
        self.worksheets[name] = self.workbook.add_worksheet(name)
        return name

    def build_table(self, rows, columns, vlookup=None):
        table = []
        if vlookup:
            for i in range(0, len(rows)):
                list_row = list(rows[i])
                lookup_column = vlookup['lookup_column'] + str(i+2) 
                xlsx_vlookup = XlsxVlookup(lookup_column, vlookup['table_start'], vlookup['table_end'], vlookup['column_index'], sheet=vlookup['sheet_name'])
                list_row.insert(vlookup['column_insert_index'], xlsx_vlookup.get_formula())
                table.append(list_row)
            columns.insert(vlookup['column_insert_index'], vlookup['column_name'])
        else:
            table = list(map(list, rows))
        table.insert(0, columns)
        self.table = table

    def write(self, sheet_name, row, column, data, format=None):
        data_len = len(str(data)) + 2 #xlsxwriter column length is not precisely string length
        current_size = data_len
        col_sizes = self.worksheets[sheet_name].col_sizes
        if column in col_sizes:
            current_size = col_sizes[column]
        length = max(data_len, current_size)
        if length == data_len:
            self.worksheets[sheet_name].set_column(column, column, length)
        self.worksheets[sheet_name].write(row, column, data, format)

    def close(self):
        self.workbook.close()

    def perform(self, column_list, rows, sheet_name=None, vlookup=None):
        sheet_name = self.create_worksheet(sheet_name)
        if vlookup:
            self.workbook.add_worksheet(vlookup['sheet_name'])
        self.build_table(rows,column_list,vlookup=vlookup)
        self.export(sheet_name)
        self.close()

    def export(self, sheet_name):
        column = True
        for i in range(0,len(self.table)):
            if i == 1:
                column = False
            for j in range(0,len(self.table[0])):
                cell = self.table[i][j]
                if column:
                    self.write(sheet_name, i, j, cell, self.column_format)
                elif self.blank_nulls and (cell == None or cell == 'None'):
                    self.write(sheet_name, i, j, ' ')
                elif self.date_format and (isinstance(cell, date) or isinstance(cell, datetime)):
                    self.write(sheet_name, i, j, cell, self.date_format)
                else:
                    self.write(sheet_name, i, j, str(cell))

#Reference another sheet in a vlookup on the export.  Experimental.
class XlsxVlookup(object):

    def __init__(self, lookup_cell, table_start, table_end, column_index, sheet=None):
        self.lookup_cell = lookup_cell
        self.table = self.build_table(table_start, table_end, sheet=sheet)
        self.column_index = column_index

    def build_table(self, table_start, table_end, sheet=None):
        if sheet:
            return "'%s'!%s:%s" % (sheet, table_start, table_end)
        else:
            return '%s:%s' % (table_start, table_end)

    def get_formula(self):
        return '=vlookup(%s,%s,%s,FALSE)' % (self.lookup_cell, self.table, self.column_index)



