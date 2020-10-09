excel_store

DISCLAIMER
----------
This is a very simple tool to load, process and query Excel-based information in an RDBMS.
I often encountered the situation that Excel data was needed in database, so I could combine/query with other data.
There are several ways to handle this, still I usually ended up writing new create tables and generating new CSV-s, sqlloaders.

This solution creates a single storage for all excels once and reuses later.
You can dump any excel here. (Well, not anything, but most of the common situations can be handled even with this first primitive version. Only columns A-Z will be loaded from each sheet.)

There is an option to generate views on the top of this generic store, to make your datasets more readable. You can then combine your views and use the excel data in queries.
Default view is generated based on the first row.
If you need a separate view for another portion of your excel, please create it manually.

USE AT YOUR OWN RISK AND RESPONSIBILITY!

Prerequsites:
-------------
1. Get Python installed
2. pip install openpyxl
3. Create the store table in RDBMS (MySql, Oracle DDL prepared, modify for other platforms)
MySQL:
	use <yourdatabase>
	source excel_store_mysql_ddl.sql

Oracle:
	@excel_store_ora_ddl.sql


Usage:
------
Invoke excel_store_gen.py from Windows cmd or cygwin console
P1 - excel file (mandatory)
P2 - sheet (optional), default is "#ALL#" which processes all sheets
P3 - view generation required (optional) default is "Y"

examples: 
py excel_store_gen.py "test1.xlsx"
	It will generate two files for each sheet.

py excel_store_gen.py "test1.xlsx" "PRODUCT"
	It will generate two files for this sheet only.
	  The first one contains the inserts (there is a delete before).
	  The second one is a create view ddl.

py excel_store_gen.py "test1.xlsx" "PRODUCT" "N"
	It will generate one file for this sheet only.
	  It contains the inserts (there is a delete before).
	  No create view ddl is generated.

Copy your files content to MySQL Workbench or PL/SQL Developer or other SQL IDE and execute.
OR execute directly from command line sql:

mysql> use <yourdatabase>
mysql> source <script(s)>

sqlplus> set escape on
sqlplus> set sqlblanklines on
sqlplus> @<script(s)>

GOOD LUCK!