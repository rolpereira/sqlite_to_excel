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

If you wish the run the tests in the folder `t` you also need to have
the following packages installed.

* File::Slurp
* Spreadsheet::Read

# Installation #

There is no installation, simply download the script
`sqlite_to_excel.pl` into your machine and run it using the Perl
executable.

## Running the tests ##

If you want to run the integration tests for the script download the
folder `t` into the same directory as the `sqlite_to_excel.pl` script
and run the `prove` command on the command line:

    rolando@rolando-K8NF4G-VSTA:~/src/git/sqlite_to_excel$ prove
    t/01-write-single-sheet.t .. ok
    All tests successful.
    Files=1, Tests=3,  3 wallclock secs ( 0.09 usr  0.02 sys +  0.88 cusr  0.12 csys =  1.11 CPU)
    Result: PASS


# Usage #

To run the query `SELECT info_number FROM data` against the database
`db.sqlite` and store it in the `newfile.xlsx` Excel file run the
following command:
    
    perl sqlite_to_excel.pl --db="./db.sqlite" --excel_path="./newfile.xlsx" --query="SELECT info_number FROM data"

# License #

This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.
