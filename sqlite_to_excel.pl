#!/usr/bin/perl

# sqlite_to_excel.pl - Extract information from a SQLite database and put it in a Excel spreadsheet
#
# License:
#
#
# Example:
#
#     perl sqlite_to_excel.pl --db="./db.sqlite" --excel_path="./newfile.xlsx" --query="SELECT info_number FROM data"
#
# Copyright 2014, Rolando Pereira, rolando_pereira@sapo.pt
#
# All Rights Reserved. This module is free software. It may be used,
# redistributed and/or modified under the same terms as Perl itself.


use v5.8.0;

use strict;
use warnings;

use Getopt::Long;

use DBI;
use DBD::SQLite;
use Excel::Writer::XLSX;

my $db_path;
my $query;
my $excel_path;

GetOptions("db=s"         => \$db_path,
           "query=s"      => \$query,
           "excel_path=s" => \$excel_path)
  or die "Error in command line arguments";

if (not $db_path) {
  usage_help();
  die "No SQLite database was given. See the flag \"--db\"";
}

if (not $query) {
  usage_help();
  die "No SQL query was given. See the flag \"--query\"";
}

if (not $excel_path) {
  usage_help();
  die "No Excel file was given. See the flag \"--excel_path\"";
}

if ( -e $excel_path ) {
  die "Excel file \"$excel_path\" already exists";
}

my $dbh = DBI->connect("dbi:SQLite:dbname=$db_path","","") or die $DBI::errstr;

my $sth = $dbh->prepare($query) or die $DBI::errstr;

$sth->execute() or die "Couldn't execute query: " . $DBI::errstr;

my $workbook = Excel::Writer::XLSX->new($excel_path);
my $worksheet = $workbook->add_worksheet();

my $row_counter = 0;

while (my $row = $sth->fetchrow_arrayref()) {
  my $column_counter = 0;

  foreach my $column (@$row) {
    $worksheet->write($row_counter, $column_counter, $column);
    $column_counter++;
  }

  $row_counter++;
}

print "Done\n";

exit 0;

sub usage_help {
  print <<"EOF";
    $0 - Run a query in a SQLite database and put its contents in a Excel spreadsheet

USAGE:

    $0 --db=<sqlite_database_path> --excel_path=<name_of_file> --query="<SQL query>"

Flags:

    --db             Path to the SQLite database

    --query          SQL query to be executed in the database

    --excel_path     Path of the Excel file that will be created with data from the SQL query

Example:

   perl $0 --db="./db.sqlite" --excel_path="./newfile.xlsx" --query="SELECT info_number FROM data"

EOF

  return;
}
