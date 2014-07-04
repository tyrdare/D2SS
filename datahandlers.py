__author__ = 'oradba'

import xlsxwriter
#import csv
import ezodf
import datetime




class DataWriter(object):
    """
    base class for db_specific query handling actions
    such as getting the query metadata (column names, types, etc.
    """

    def __init__(self, curs, db_type, output_type, output_file, write_header):
        """
        Takes a cursor to instantiate
        """
        self.curs = curs
        self.db_type = db_type
        self.column = 0
        self.row = 0
        self.write_header = write_header
        self.output_type = self.check_output_type(output_type)
        self.header = self.set_header()
        self.output_file = output_file
        self.output_dest = self.set_output_dest()

        # need to find out what sort of database I'm getting data from.
        # use the column type info from the cursor's description attributes
        # to determine the type of the output column (string, number, date, etc.

    def check_output_type(self,output_type):
        output_types = ["CSV", "XLSX", "ODS"]
        #output_types = ["XLSX"]
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
            spreadsheet = xlsxwriter.Workbook(self.output_file)
            spreadsheet.add_worksheet("Data")
            # need to create this and save it for writing dates later.
            self.date_format = spreadsheet.add_format({'num_format': 'yyyy/mm/dd hh:mm:ss'})


        #elif self.output_type == "CSV":
        #    spreadsheet = csv.writer(self.output_file,delimiter=":::")

        elif self.output_type == "ODS":
            spreadsheet = ezodf.newdoc(doctype="ods", filename=self.output_file)
            spreadsheet.sheets.append(ezodf.Sheet("Data"))
            spreadsheet.save()

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

        if self.output_type == 'ODS':
            self.output_dest.save

    def write_header_row(self):
        """
        Writes out the header row. All Columns are assumed to be strings
        """

        self.write_row(self.header)


    def write_row(self, row):
        if self.output_type == "XLSX":
            #print "write_row(): in the XLS block"
            #print "data:", row
            for i in range(len(row)):
                print row[i], "is type", type(row[i])
                # first condition catches Oracle dates, second catches PostgreSQL dates for XLSX output
                # XLSX will treat outputted python datetimes as numbers otherwise
                if type(row[i]) == datetime.datetime or type(row[i]) == datetime.date:
                    self.output_dest.worksheets()[0].write_datetime(self.row, self.column + i, row[i], self.date_format)

                else:
                    self.output_dest.worksheets()[0].write(self.row, self.column+i, row[i] )
            self.row += 1

        elif self.output_type == "ODS":
            #print "write_row(): in the ODS block"
            #print "data:", row
            for i in range(len(row)):
                self.output_dest.sheets[0][self.row, self.column + i].set_value(row[i])
            self.row += 1

        elif self.output_type == "CSV":
            pass


    def close(self):
        if self.output_type == 'XLSX':
            # xlsxwriter doesn't close its file, it just exits.
            self.output_dest.close()

        elif self.output_type == 'ODS':
            self.output_dest.save()
        elif self.output_type == 'CSV':
            self.output_dest.close()