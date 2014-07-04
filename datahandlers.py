__author__ = 'oradba'

import cx_Oracle
import xlsxwriter
#import csv
#import ezodf
import datetime


class ColumnDescription(object):
    def __init__(self):
        pass

# data handlers differ by database and by the type definition of the columns returned.

class DataWriter(object):
    """
    base class for db_specific query handling actions
    such as getting the query metadata (column names, types, etc.
    """

    def __init__(self, curs, db_type, output_type, output_file):
        """
        Takes a cursor to instantiate
        """
        self.curs = curs
        self.db_type = db_type
        self.column = 0
        self.row = 0
        self.output_type = self.check_output_type(output_type)
        self.header = self.set_header()
        self.output_file = output_file
        self.output_dest = self.set_output_dest()

        # need to find out what sort of database I'm getting data from.
        # use the column type info from the cursor's description attributes
        # to determine the type of the output column (string, number, date, etc.

    def check_output_type(self,output_type):
        #output_types = ["CSV", "XLSX", "ODS"]
        output_types = ["XLSX"]
        if output_type in output_types:
            return output_type
        else:
            raise TypeError("Invalid output type, must be one of %s" % ",".join(output_types))

    def set_output_dest(self):
        """
        Creates the XLSX, ODS spreadsheet or CSV file and returns a reference to it.
        """
        spreadsheet = None

        if self.output_type == "XLSX":
            # Can't save a Workbook - can only close()
            spreadsheet =  xlsxwriter.Workbook(self.output_file)
            spreadsheet.add_worksheet("Data")
            # need to create this and save it for writing dates later.
            self.date_format = spreadsheet.add_format({'num_format': 'yyyy/mm/dd hh:mm:ss'})


        #elif self.output_type == "CSV":
        #    spreadsheet = csv.writer(self.output_file,delimiter=":::")

        #elif self.output_type == "ODS":
        #    spreadsheet = ezodf.newdoc(doctype="ods", filename=self.output_file)
        #    spreadsheet.sheets[0] = ezodf.Sheet("Data")
        #    spreadsheet.save()

        return spreadsheet

    def set_header(self):
        """
        Creates a column names list from the cursor.description attribute supported by DBAPI 2.0
        """
        return [x[0] for x in self.curs.description]

    def write_data(self):
        if self.write_header:
            self.write_header_row()

        for row in self.curs:
            self.write_row(row)

    def write_header_row(self):
        # need to know what format I'm writing out to: CSV, ODF or  XLS

        if self.output_type == "XLSX":
            #print "write_header(): in the XLS block"
            self.write_row(self.header)
        #elif self.output_type == "CSV":
        #    pass
        #elif self.output_type == "ODS":
        #    pass

    def write_row(self, row):
        if self.output_type == "XLSX":
            #print "write_row(): in the XLS block"
            #print "data:", row
            for i in range(len(row)):
                if type(row[i]) == datetime.datetime:
                    self.output_dest.worksheets()[0].write_datetime(self.row,self.column + i, row[i], self.date_format)
                else:
                    self.output_dest.worksheets()[0].write(self.row, self.column+i, row[i] )
            self.row +=1
        #elif self.output_type == "CSV":
        #    pass
        #elif self.output_type == "ODS":
        #    pass

