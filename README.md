Run a query against a SQLite database and store the result in a Excel
work sheet.

# Dependencies #

You need to have Perl v5.8.0 installed.

To run this script you also need to have the following packages
installed from CPAN.

* DBI
* DBD::SQLite
* Excel::Writer::XLSX

There is no need to have Microsoft Excel or LibreOffice installed on
the machine that runs this script.

# Installation #

There is no installation, simply download the script
`sqlite_to_excel.pl` into your machine and run it using the Perl
executable.

# Usage #

To run the query `SELECT info_number FROM data` against the database
`db.sqlite` and store it in the `newfile.xlsx` Excel file run the
following command:
    
    perl sqlite_to_excel.pl --db="./db.sqlite" --excel_path="./newfile.xlsx" --query="SELECT info_number FROM data"

# License #

This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.
