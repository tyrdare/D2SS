D2SS
======

A utility to write the data from an arbitrary query to an spreadsheet.

In order to use D2SS you have to have the following dependencies installed:

:OS-Level packages/software:
    Oracle instant client

    PostgreSQL client library



:Python dependencies:
    psycopg2

    cx_Oracle

    sqlite3

    xlsxwriter

    ezodf

    csv

Currently supported:
""""""""""""""""""""
:Databases:

    Oracle

:Output formats:
    XLSX

Usage
"""""
::

[python] d2ss.py [-h] [-f FILE] [-ld] [-lo]
 
A Program to output query data to a spreadsheet (XLS or CSV)

optional arguments:
   -h, --help  show this help message and exit
   -f FILE     path to the JSON-formatted configuration file
   -ld         List program-supported databases
   -lo         List program-supported output formats
 

Configuration
"""""""""""""
The configuration file is a simple text file containing a javascript object in JSON.

::
 
{

    "db": {

        "_comment": "for connect_string, use a string appropriate to your database flavor. db_type is the database flavor: ORCL, PGSQL, MYSQL, MSSQL, SQLITE",

        "connect_string":"database_specific_connect_string",

        "db_type": "ORCL"

    },

    "output": {

        "_comment": "output_headers will put column headers on the column. output_type can be CSV, XLS, ODS",

        "output_headers": true,

        "output_type": "XLSX",

        "output_file": "/path/to/your/output.file"

    },

    "_comment": "query is an array of clauses that make up the sql statement.  These will be concatenated in the program but allow you to make the statement a little more readable",

    "query" :

        [

            "select \*",

            "from some_table",

            "where condition",

            "order by column set"

        ]

}


Note that the query is in multiple elements of an array.  This is so the user can format the sql for readability.  The program will glue it all back together into a string using spaces to execute it.
