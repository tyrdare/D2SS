__author__ = 'oradba'

import os
import xlsxwriter
import csv
import ezodf
import datetime


class DataWriter(object):
    """
    base class for db_specific query handling actions
    such as getting the query metadata (column names, types, etc.
    """

    def __init__(self, curs, db_type, output_type, output_file, write_header):
        """
        cursor: a database cursor
        db_type: string, description of the database module type
        output_type, string, XLSX, CSV or ODS
        output_file:  string, path and name of the output file
        write_header:  bool, flag indicating if columns headers should be output.
        """
        self.curs = curs
        self.db_type = db_type
        self.column = 0
        self.row = 0
        self.write_header = write_header
        self.date_format = None
        self.output_type = self.check_output_type(output_type)
        self.header = self.set_header(curs)

        if not os.path.exists(os.path.split(output_file)[0]):
            raise OSError("Path %s does not exist" % os.path.split(output_file)[0])
        else:
            self.output_file = output_file

        self.output_dest = self.set_output_dest(self.output_type, self.output_file)
        if self.output_type == 'XLS':
            # need to create this and save it for writing dates later.
            self.date_format = self.output.dest.add_format({'num_format': 'yyyy/mm/dd hh:mm:ss'})


    @staticmethod
    def check_output_type(output_type):
        output_types = ["CSV", "XLSX", "ODS"]
        #output_types = ["XLSX"]
        if output_type in output_types:
            return output_type
        else:
            raise TypeError("Invalid output type, must be one of %s" % ",".join(output_types))

    @staticmethod
    def set_output_dest(output_type, output_file):
        """
        Creates the XLSX, ODS spreadsheet or CSV file and returns a reference to it.
        """
        spreadsheet = None

        if output_type == "XLSX":
            # Can't save a Workbook - can only close()
            spreadsheet = xlsxwriter.Workbook(output_file)
            spreadsheet.add_worksheet("Data")

        elif output_type == "CSV":

            f = open(output_file, "w")
            spreadsheet = csv.writer(
                f, dialect="excel", delimiter="~", quoting=csv.QUOTE_NONNUMERIC, escapechar='^', doublequote=False
            )

        elif output_type == "ODS":

            spreadsheet = ezodf.newdoc(doctype="ods", filename=output_file)
            # add a sheet to the empty sheets list
            spreadsheet.sheets.append(ezodf.Sheet("Data"))
            spreadsheet.save()

        return spreadsheet

    @staticmethod
    def set_header(curs):
        """
        Creates a column names list from the cursor.description attribute supported by DBAPI 2.0
        """
        return [x[0] for x in curs.description]

    def write_data(self):
        """
        Writes the header if self.write_header is True then proceeds to write the rows
        """
        if self.write_header:
            self.write_header_row()

        for row in self.curs:
            self.write_row(row)


    def write_header_row(self):
        """
        Writes out the header row. All Columns are assumed to be strings
        """

        self.write_row(self.header)

    def write_row(self, row):
        if self.output_type == "XLSX":

            for i in range(len(row)):

                # first condition formats Oracle dates and PostgreSQL timestamps,
                # second catches PostgreSQL dates for XLSX output
                # XLSX will treat outputted python datetimes as numbers otherwise
                if type(row[i]) == datetime.datetime or type(row[i]) == datetime.date:
                    self.output_dest.worksheets()[0].write_datetime(self.row, self.column + i, row[i], self.date_format)
                else:
                    self.output_dest.worksheets()[0].write(self.row, self.column+i, row[i])
            self.row += 1

        elif self.output_type == "ODS":

            for i in range(len(row)):
                self.output_dest.sheets[0][self.row, self.column + i].set_value(row[i])
            self.row += 1

        elif self.output_type == "CSV":
            self.output_dest.writerow(row)

    def close(self):
        if self.output_type == 'XLSX':
            # xlsxwriter doesn't save its file, it just closes
            self.output_dest.close()

        elif self.output_type == 'ODS':
            # ezodf doesn't close files, it just saves
            self.output_dest.save()

        elif self.output_type == 'CSV':
            # and we just have to go out of scope for csv to close a file.
            pass