__author__ = 'oradba'


import sys
import argparse
import config

import datahandlers


connection = None
db_module = None

# Types of databases this program will support
db_flavors = {
    "ORCL": "cx_Oracle",
    "PGSQL": "psycopg2",
    "MYSQL": "mysqldb",
    "MSSQL": "pyodbc"
}

output_flavors = {
    "XLSX": "Excel Spreadheet",
    "CSV": "Comma Separated Values",
    "ODS": "Open Document Spreadsheet"
}


class ConfigurationError(Exception):
    pass


def show_supported_dbs():
    """
    Display the list of supported databases
    """
    global db_flavors

    print "Supported databases are:"
    for key in db_flavors.keys():
        print "\t%s : %s" % (key, db_flavors[key])


def show_supported_output_formats():
    """
    Display the list of supported output formats
    """
    global output_flavors

    print "Supported output formats:"
    for key in output_flavors.keys():
        print "\t%s : %s" % (key, output_flavors[key])


def set_database_flavor(flavor):
    """
    Determines the database type and attempts to dynamically import the correct module for it.
    """
    global db_module, db_flavors

    #mesg = "Database type is %s"
    # if the requested database type is one we support, try to import the driver
    if flavor.upper() in db_flavors:
        db_module = __import__(db_flavors[flavor.upper()])
    else:
        print "Database type %s is not supported" % flavor
        show_supported_dbs()


def get_db_connection(connect_string):
    """
    Attempts to open a db connection and return same.
    """
    global connection, db_module

    if connect_string:
        try:
            connection = db_module.connect(config.db_connect_string)
        except db_module.DatabaseError as e:
            err, = e.args
            print "Error attempting to get a %s connection: %s" % (db_module.__name__, err)
            sys.exit(3)

        return connection
    raise ConfigurationError("No database connection configured")

def process_args():
    """
    Returns an args object
    """
    arg_obj = argparse.ArgumentParser(description="A Program to output query data to a spreadsheet (XLS, ODS or CSV)")
    arg_obj.add_argument("-ld", action="store_true", help="List program-supported databases")
    arg_obj.add_argument("-lo", action="store_true", help="List program-supported output formats")
    return arg_obj.parse_args()


def execute_query(query):
    """
    Execute the submitted query and return the cursor
    query param is a list
    """

    # we won't need to keep the database connection, since the cursor will hold a reference to it
    conn = get_db_connection(config.db_connect_string)
    try:
        curs = conn.cursor()
    except db_module.DatabaseError as dbe:
        err, = dbe.args
        if hasattr(err, "message"):
            print "Error obtaining database cursor: %s" % err.message
        else:
            print "Error obtaining database cursor: %s" % err
        sys.exit(4)

    try:
        curs.execute(" ".join(query))
    except db_module.DatabaseError as dbe:
        err, = dbe.args
        if hasattr(err, "message"):
            print "Error executing SQL statement: %s" % err.message
        else:
            print "Error executing SQL statement: %s" % err
        sys.exit(4)
    except db_module.DataError as de:
        err, = de.args
        if hasattr(err, "message"):
            print "Error executing SQL statement: %s" % err.message
        else:
            print "Error executing SQL statement: %s" % err
        sys.exit(5)

    return curs


def get_data_writer(curs, db_type, output_type, output_file, output_headers):

    if output_type == "XLSX":
        #import xlswriter
        writer = datahandlers.XlsxDataWriter(curs, db_type, output_type, output_file, output_headers)
    elif output_type == "CSV":
        #import csv
        writer = datahandlers.CsvDataWriter(curs, db_type, output_type, output_file, output_headers)
    elif output_type == "ODS":
        #import ezodf
        writer = datahandlers.OdsDataWriter(curs, db_type, output_type, output_file, output_headers)
    else:
        raise TypeError("Unsupported output type: %s" % output_type)
    return writer


def main():
    """
    Make main routine importable.  Important in some situations
    """
    exit_flag = False
    # get my command line arguments
    arg_obj = process_args()

    if arg_obj.ld:
        show_supported_dbs()
        exit_flag = True
    if arg_obj.lo:
        show_supported_output_formats()
        exit_flag = True

    if exit_flag:
        sys.exit(0)

    set_database_flavor(config.db_type)

    curs = execute_query(config.query)

    # gimme an xlsx DataWriter
    # DataWriter is responsible for creating the spreadsheet
    dw = get_data_writer(curs, config.db_type, config.output_type, config.output_file, config.output_headers)
    dw.write_data()
    dw.close()
    conn = curs.connection
    curs.close()
    conn.close()


if __name__ == "__main__":
    main()
