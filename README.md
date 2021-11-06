# Reports
Python scripts used to automate reports. The basic structure is:
 1) Send a PostgreSQL query throught psycopg2 (Some are conformed dynamically within the scripts)
 2) Process or reformat the returned tuple.
 3) Push the content to an excel sheet using OpenPyXl.
 4) Format the cells accordingly.
 
Modules used:
- OpenPyXl
- Psycopg2
