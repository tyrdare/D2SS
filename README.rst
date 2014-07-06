D2SS
====

Database To SpreadSheet

A utility to write the data from an arbitrary query to an spreadsheet.

In order to use D2SS you have to have the following dependencies installed:

:OS-Level packages/software:
    Oracle instant client

    PostgreSQL client library

    mysqldb or mariadb client library



:Python dependencies:
    psycopg2

    cx_Oracle

    mysqldb-python

    sqlite3

    xlsxwriter

    ezodf

    csv

These python dependencies are loaded when needed based on the db type and output type requested.  If you're only
going to use Oracle databases and Excel spreadsheets, you only need to install cx_Oracle and xlsxwriter on your
system.

Currently supported:
""""""""""""""""""""
Databases:

    Oracle

    PostgreSQL

Output formats:

    XLSX

    ODS

    CSV

TODO
""""
Add Database Support

    SQLite

    MSSQL

    MySQL

Usage
"""""
::

[python] d2ss.py [-h] [-ld] [-lo]
 
A Program to output query data to a spreadsheet (XLS or CSV)

optional arguments:
   -h, --help  show this help message and exit
   -ld         List program-supported databases
   -lo         List program-supported output formats
 

Configuration
"""""""""""""
The configuration file is a simple python file, config.py, containing a a set of variable and values.  It's simply
imported into the d2ss program. Please have a look

Note that "query" is an array of strings.  This is so the user can format the sql for readability.  The program will
glue it all back together into a string using spaces to execute it.
