

# for connect_string, use a string appropriate to your database flavor. db_type is the
db_connect_string = "dbname=d2ss host=127.0.0.1 port=5432 user=d2ss password=d2ss"
# database flavor: ORCL, PGSQL, MYSQL, MSSQL
db_type = "PGSQL"

# output_headers will put column headers in the first row of the file.
output_headers = True

# output_type can be CSV, XLSX, ODS
output_type = "CSV"

output_file = "/tmp/d2ss_test_pgsql.csv"


# query is an array of clauses that make up the sql statement.  These will be concatenated in the program but
# allow you to make the statement a little more readable"
query = [
    "select *",
    "from d2ss.test_table",
    "where a_float > 2",
    "order by a_float"
]
