__author__ = 'oradba'

import os
import sys
import datetime
import importlib
import csv



class OdsDataWriter(object):
    """
    Reads data from a database cursor and writes it to a Open Document Spreadsheet
    """
    io_mod = __import__('ezodf')
    def __init__(self, curs, db_type, output_type, output_file, write_header):
        """
        Set some necessary variables
        """
        #self.io_mod = None
        #self.import_mod()
        self.curs = curs
        self.db_type = db_type
        self.column = 0
        self.row = 0
        self.write_header = write_header
        self.date_format = None
        self.output_type = self.check_output_type(output_type)
        self.header = self.set_header(curs)

        self.check_output_path(output_file)

        self.output_dest = self.set_output_dest(self.output_type, self.output_file)

    def check_output_path(self, output_file):
        if not os.path.exists(os.path.split(output_file)[0]):
            raise OSError("Path %s does not exist" % os.path.split(output_file)[0])
        self.output_file = output_file

    def check_output_type(self, output_type):
        if output_type != 'ODS':
            raise TypeError("incorrect writer (ODS) for type %s" % output_type)
        return output_type

    def set_header(self, curs):
        """
        Creates a column names list from the cursor.description attribute supported by DBAPI 2.0
        """
        return [x[0] for x in curs.description]

    def set_output_dest(self, output_type, output_file):
        """
        Instantiates the spreadsheet
        """
        spreadsheet = None
        #print "In OdsWriter"
        spreadsheet = self.io_mod.newdoc(doctype="ods", filename=output_file)
        # add a sheet to the empty sheets list
        spreadsheet.sheets.append(self.io_mod.Sheet("Data"))
        print spreadsheet.sheets
        spreadsheet.save()
        return spreadsheet

    def write_data(self):
        """
        Writes all the data to the spreadsheet
        """
        if self.write_header:
            self.write_header_row()

        for row in self.curs:
            self.write_row(row)

    def write_row(self, row):
        #print "number of sheets is ", len(self.output_dest.sheets)
        for i in range(len(row)):
            self.output_dest.sheets[0][self.row, self.column + i].set_value(row[i])
        self.row += 1

    # Overrides superclass' method
    def write_header_row(self):
        self.write_row(self.header)

    def close(self):
        self.output_dest.save()


class XlsxDataWriter(object):
    io_mod = __import__('xlsxwriter')

    def __init__(self, curs, db_type, output_type, output_file, write_header):
        """
        Set some necessary variables
        """
        import xlsxwriter
        self.curs = curs
        self.db_type = db_type
        self.column = 0
        self.row = 0
        self.write_header = write_header
        self.date_format = None
        self.output_type = self.check_output_type(output_type)
        self.header = self.set_header(curs)

        self.check_output_path(output_file)

        self.output_dest = self.set_output_dest(self.output_type, self.output_file)
        if self.output_type == 'XLS':
            # need to create this and save it for writing dates later.
            self.date_format = self.output.dest.add_format({'num_format': 'yyyy/mm/dd hh:mm:ss'})
        #print self.__dict__
        #sys.exit(0)

    def check_output_path(self, output_file):
        if not os.path.exists(os.path.split(output_file)[0]):
            raise OSError("Path %s does not exist" % os.path.split(output_file)[0])
        self.output_file = output_file

    def check_output_type(self, output_type):
        if output_type != 'XLSX':
            raise TypeError("incorrect writer (ODS) for type %s" % output_type)
        return output_type

    def set_header(self, curs):
        """
        Creates a column names list from the cursor.description attribute supported by DBAPI 2.0
        """
        return [x[0] for x in curs.description]

    def set_output_dest(self, output_type, output_file):
        """
        Instantiates the spreadsheet
        """
        spreadsheet = None
        spreadsheet= self.io_mod.Workbook(output_file)
        spreadsheet.add_worksheet("Data")
        return spreadsheet

    def write_data(self):
        """
        Writes all the data to the spreadsheet
        """
        if self.write_header:
            self.write_header_row()

        for row in self.curs:
            self.write_row(row)

    def write_row(self, row):
        #print "number of sheets is ", len(self.output_dest.sheets)
        for i in range(len(row)):
            # first condition formats Oracle dates and PostgreSQL timestamps,
            # second catches PostgreSQL dates for XLSX output
            # XLSX will treat outputted python datetimes as numbers otherwise
            if type(row[i]) == datetime.datetime or type(row[i]) == datetime.date:
                self.output_dest.worksheets()[0].write_datetime(self.row, self.column + i, row[i], self.date_format)
            else:
                self.output_dest.worksheets()[0].write(self.row, self.column+i, row[i])
        self.row += 1

    # Overrides superclass' method
    def write_header_row(self):
        self.write_row(self.header)

    def close(self):
        self.output_dest.close()

class CsvDataWriter(object):
    """
    Reads data from a database cursor and writes it to a Open Document Spreadsheet
    """
    io_mod = __import__('csv')
    def __init__(self, curs, db_type, output_type, output_file, write_header):
        """
        Set some necessary variables
        """
        import csv
        #self.io_mod = None
        #self.import_mod()
        self.curs = curs
        self.db_type = db_type
        self.column = 0
        self.row = 0
        self.write_header = write_header
        self.date_format = None
        self.output_type = self.check_output_type(output_type)
        self.header = self.set_header(curs)

        self.check_output_path(output_file)

        self.output_dest = self.set_output_dest(self.output_type, self.output_file)

    def check_output_path(self, output_file):
        if not os.path.exists(os.path.split(output_file)[0]):
            raise OSError("Path %s does not exist" % os.path.split(output_file)[0])
        self.output_file = output_file

    def check_output_type(self, output_type):
        if output_type != 'CSV':
            raise TypeError("incorrect writer (CSV) for type %s" % output_type)
        return output_type

    def set_header(self, curs):
        """
        Creates a column names list from the cursor.description attribute supported by DBAPI 2.0
        """
        return [x[0] for x in curs.description]

    def set_output_dest(self, output_type, output_file):
        """
        Instantiates the output file
        """
        spreadsheet = None
        f = open(output_file, "w")
        spreadsheet = self.io_mod.writer(f, dialect="excel", delimiter="~", quoting=csv.QUOTE_NONNUMERIC, escapechar='^', doublequote=False)
        return spreadsheet

    def write_data(self):
        """
        Writes all the data to the spreadsheet
        """
        if self.write_header:
            self.write_header_row()

        for row in self.curs:
            self.write_row(row)

    def write_row(self, row):
        #print "number of sheets is ", len(self.output_dest.sheets)
        self.output_dest.writerow(row)

    # Overrides superclass' method
    def write_header_row(self):
        self.write_row(self.header)

    def close(self):
        pass