__author__ = 'oradba'


import sys
import os
import argparse
import config

from datahandlers import DataWriter


connection = None
db_module = None
db_type = None
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
    global db_flavors

    print "Supported databases are %s" % ",".join(db_flavors.keys())
    sys.exit(0)


def show_supported_output_formats():
    print "Supported output formats:"
    for key in output_flavors.keys():
        print "%s : %s" % (key, output_flavors[key])
        sys.exit(0)


def set_database_flavor(flavor):
    """
    Determines the database type and attempts to import the correct module for it.
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
    arg_obj = argparse.ArgumentParser(description="A Program to output query data to a spreadsheet (XLS or CSV)")
    arg_obj.add_argument("-ld", action="store_true", help="List program-supported databases")
    arg_obj.add_argument("-lo", action="store_true", help="List program-supported output formats")
    return arg_obj.parse_args()


def execute_query(query):

    # print config['db']["connect_string"]
    conn = get_db_connection(config.db_connect_string)
    try:
        curs = conn.cursor()
    except db_module.DatabaseError as dbe:
        err, = dbe.args
        if hassattr(err, "message"):
            print "Error obtaining database cursor: %s" % err.message
        else:
            print "Error obtaining database cursor: %s" % err
        sys.exit(4)

    try:
        curs.execute(" ".join(config.query))
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


def main():
    # get my command line arguments
    arg_obj = process_args()
    #print arg_obj
    if arg_obj.ld:
        show_supported_dbs()
    if arg_obj.lo:
        show_supported_output_formats()

    set_database_flavor(config.db_type)

    curs = execute_query(config.query)

    dw = DataWriter(curs, config.db_type, config.output_type, config.output_file, config.output_headers)
    dw.write_data()

    conn = curs.connection
    curs.close()
    conn.close()


if __name__ == "__main__":
    main()
