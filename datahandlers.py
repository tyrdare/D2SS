__author__ = 'oradba'

import os
import sys
import datetime
import importlib
import csv
from abc import ABCMeta, abstractmethod


class DataWriter(object):

    __meta__ = ABCMeta
    output_type = None


    def __init__(self, curs, db_type, output_type, output_file, write_header):
        """
        Set some necessary variables
        """
        self.curs = curs
        self.db_type = db_type
        self.column = 0
        self.row = 0
        self.write_header = write_header
        self.date_format = None
        self.output_type = self.check_output_type(output_type)
        self.header = self.set_header(curs)
        self.check_output_path(output_file)

        self.output_file = self.check_output_path(output_file)
        self.output_dest = self.set_output_dest(self.output_type, self.output_file)


    def check_output_path(self, output_file):
        if not os.path.exists(os.path.split(output_file)[0]):
            raise OSError("Path %s does not exist" % os.path.split(output_file)[0])
        return output_file

    def check_output_type(self, output_type):
        if output_type != self.output_type:
            raise TypeError("incorrect writer (%s) for type %s" % (self.output_type,output_type))
        return output_type

    @staticmethod
    def set_header(curs):
        """
        Creates a column names list from the cursor.description attribute supported by DBAPI 2.0
        """
        return [x[0] for x in curs.description]

    @abstractmethod
    def set_output_dest(self):
        """
            return a an output destination resource
        """
        pass

    @abstractmethod
    def write_data(self):
        pass

    @abstractmethod
    def write_row(self):
        pass

    @abstractmethod
    def write_header_row(self):
        pass

    @abstractmethod
    def close(self):
        pass

#===============================
class OdsDataWriter(DataWriter):
    """
    Reads data from a database cursor and writes it to a Open Document Spreadsheet
    """
    output_type = 'ODS'
    io_mod = __import__('ezodf')


    def set_output_dest(self, output_type, output_file):
        """
        Instantiates the spreadsheet
        """
        spreadsheet = None
        spreadsheet = self.io_mod.newdoc(doctype="ods", filename=output_file)
        # add a sheet to the empty sheets list
        spreadsheet.sheets.append(self.io_mod.Sheet("Data"))
        spreadsheet.save()
        return spreadsheet

    def write_data(self):
        """
        Writes all the data to the spreadsheet
        """
        if self.write_header:
            self.write_header_row()
        [self.write_row(row) for row in self.curs]


    def write_row(self, row):
        #print "number of sheets is ", len(self.output_dest.sheets)
        #for i in range(len(row)):
        #    self.output_dest.sheets[0][self.row, self.column + i].set_value(row[i])
        #self.row += 1

        for i,elem in enumerate(row):
            self.output_dest.sheets[0][self.row, self.column + i].set_value(elem)
        self.row += 1

    # Overrides superclass' method
    def write_header_row(self):
        self.write_row(self.header)

    def close(self):
        self.output_dest.save()

#================================
class XlsxDataWriter(DataWriter):

    output_type= "XLSX"
    io_mod = __import__('xlsxwriter')

    def __init__(self, curs, db_type, output_type, output_file, write_header):
        """
        Set some necessary variables
        """
        super(XlsxDataWriter,self).__init__(curs, db_type, output_type, output_file, write_header)
        # need to create this and save it for writing dates later.
        self.date_format = self.output_dest.add_format({'num_format': 'yyyy/mm/dd hh:mm:ss'})


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

        [self.write_row(row) for row in self.curs]

        #for row in self.curs:
        #    self.write_row(row)

    def write_row(self, row):
        #print "number of sheets is ", len(self.output_dest.sheets)
        for i, elem in enumerate(row):
            # first condition formats Oracle dates and PostgreSQL timestamps,
            # second catches PostgreSQL dates for XLSX output
            # XLSX will treat outputted python datetimes as numbers otherwise
            if type(elem) == datetime.datetime or type(elem) == datetime.date:
                self.output_dest.worksheets()[0].write_datetime(self.row, self.column + i, elem, self.date_format)
            else:
                self.output_dest.worksheets()[0].write(self.row, self.column+i, elem)
        self.row += 1

    # Overrides superclass' method
    def write_header_row(self):
        self.write_row(self.header)

    def close(self):
        self.output_dest.close()

#===============================
class CsvDataWriter(DataWriter):
    """
    Reads data from a database cursor and writes it to a comma separated values file
    """
    output_type = 'CSV'
    io_mod = __import__('csv')


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

        #for row in self.curs:
        #    self.write_row(row)
        [self.write_row(row) for row in self.curs]

    def write_row(self, row):
        #print "number of sheets is ", len(self.output_dest.sheets)
        self.output_dest.writerow(row)

    # Overrides superclass' method
    def write_header_row(self):
        self.write_row(self.header)

    def close(self):
        pass


if __name__ == "__main__":
    pass