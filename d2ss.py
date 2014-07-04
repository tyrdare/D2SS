__author__ = 'oradba'

#import cx_Oracle
import sys
import os
import argparse
import json

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


def show_supported_dbs():
    global db_flavors

    print "Supported databases are %s" % ",".join(db_flavors.keys())
    sys.exit(0)


def show_supported_output_formats():
    print "Supported output formats:"
    for key in output_flavors.keys():
        print "%s : %s" % (key, output_flavors[key])
        sys.exit(0)


def set_database_flavor(cfg_obj):
    """
    Determines the database type and attempts to import the correct module for it.
    """
    global db_module, db_flavors

    flavor = cfg_obj['db']["db_type"]
    #mesg = "Database type is %s"
    # if the requested database type is one we support, try to import the driver
    if flavor.upper() in db_flavors.keys():
        db_module = __import__(db_flavors[flavor.upper()])
    else:
        print "Database type %s is not supported" % flavor
        print "Supported databases are %s" % ",".join(db_flavors.keys())
        sys.exit(1)


def get_db_connection(cfg_obj):
    """
    Attempts to open a db connection and return same.
    """
    global connection, db_module
    conn_str = cfg_obj['db']["connect_string"]
    try:
        connection = db_module.connect(conn_str)
    except Exception as e:
        print "Error attempting to get a %s connection" % db_module.__name__
        sys.exit(3)

    return connection


def process_args():
    """
    Returns an args object
    """
    arg_obj = argparse.ArgumentParser(description="A Program to output query data to a spreadsheet (XLS or CSV)")
    arg_obj.add_argument(
        "-f", action="store", dest="file", type=str, help="path to the JSON-formatted configuration file"
    )
    arg_obj.add_argument("-ld", action="store_true", help="List program-supported databases")
    arg_obj.add_argument("-lo", action="store_true", help="List program-supported output formats")
    return arg_obj.parse_args()


def load_config(fname):
    # open the config file, if it exists and is a file
    if not os.path.exists(fname) or not os.path.isfile(fname):
        raise OSError("%s does not exist or is not a file" % fname)

    with open(fname, "r") as f:
        return json.load(f)


def execute_query(cfg_obj):
    set_database_flavor(cfg_obj)
    # print config['db']["connect_string"]
    conn = get_db_connection(cfg_obj)
    curs = conn.cursor()
    curs.execute(" ".join(cfg_obj["query"]))
    return curs


def main():
    # get my command line arguments
    arg_obj = process_args()
    #print arg_obj
    if arg_obj.ld:
        show_supported_dbs()
    if arg_obj.lo:
        show_supported_output_formats()

    config = load_config(arg_obj.file)

    curs = execute_query(config)

    dw = DataWriter(curs, config["db"]["db_type"], config["output"]["output_type"],  config["output"]["output_file"])
    dw.write_header = True
    dw.write_data()

    curs.connection.close()
    curs.close()

if __name__ == "__main__":
    main()
