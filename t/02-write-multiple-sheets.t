#!/usr/bin/perl

use strict;
use warnings;
use v5.8.0;

# Core perl modules
use File::Spec;
use File::Temp qw( tempfile tmpnam );
use FindBin qw( $RealBin );
use Test::More;

# Cpan modules
use DBI;
use DBD::SQLite;
use File::Slurp;
use Spreadsheet::Read;

# This test has the following steps:
#
#     - Create a SQLite database containg some data
#
#     - Call the command line script to run two queries to select that
#     data and store it in a Excel spreadsheet
#
#     - Check if the contents of the Excel spreadsheet are correct

# Prepare the database
my ($temp_fh, $temp_filename) = tempfile();

my $dbh = DBI->connect("dbi:SQLite:dbname=$temp_filename","","") or BAIL_OUT( $DBI::errstr );
####

# Parse the data file
my $data_file = File::Spec->catfile( $RealBin, "data", "01-write-single-sheet.csv");

if ( not -f $data_file ) {
  BAIL_OUT( "Cannot find file $data_file");
}

my @data_file_contents = grep { $_ !~ /^\s*$/ } # Remove empty lines
                         grep { $_ !~ /^\s*#/ } # Remove the comments from the file
                         read_file( $data_file );

# Removing trailing newline from @data_file_contents
foreach my $line ( @data_file_contents ) {
  $line =~ s/\n$//;
}

# The first line of the data file contains the header names
my @headers = split( /,/, $data_file_contents[0] );

# The remaining lines contain the data
my @data;

foreach my $data_row ( @data_file_contents[1..$#data_file_contents] ) {
  push @data, [ split( /,/, $data_row ) ];
}
####

# Inject the data into the database
my $sth_str = sprintf( "CREATE TABLE data ( %s )", join( ", ", @headers) );

$dbh->do($sth_str);

my $insert_sth_str = sprintf("INSERT INTO data(%s) VALUES ( %s )",
  join( ", ", @headers ),
  join( ", ", map { '?' } @headers ) # Add one '?' to the statement for each header
);

my $insert_sth = $dbh->prepare( $insert_sth_str );
foreach my $row ( @data ) {
  my $column_counter = 1;

  foreach my $column ( @$row ) {
    $insert_sth->bind_param( $column_counter, $column );
    $column_counter++;
  }

  $insert_sth->execute();
}

####

# Prepare a temporary excel file
my $excel_filename = tmpnam() . '.xlsx';

####

# Run the command

my $script_path = File::Spec->catfile( $RealBin, File::Spec->updir, "sqlite_to_excel.pl");

my $command = "perl $script_path --db='$temp_filename' --query='SELECT * FROM data' --query='SELECT $headers[0] FROM data' --excel='$excel_filename'";
system( $command );
####

if ( -f $excel_filename ) {
  pass( "Excel file was successfully created" );
}
else {
  BAIL_OUT( "Excel file was not created. Error: $!" );
}

# Read the sheet
my $excel_data = Spreadsheet::Read::ReadData( $excel_filename );

is( $excel_data->[0]->{sheets}, 2, 'there are two sheets in the excel file' );

my $sheet = $excel_data->[1];

is_deeply( [ Spreadsheet::Read::rows( $sheet ) ], [ @data ], 'data in first spreadsheet of the excel file matches data in database');

my $second_sheet = $excel_data->[2];

is_deeply( [ Spreadsheet::Read::rows( $second_sheet ) ],
           [ map { [ $_->[0] ] } @data ], # The "$_->[0]" corresponds to the data under the header "$header[0]"
                                          # which is the data that should be return by the second "--query" flag
           'data in second spreadsheet of the excel file matches the data in the database');

done_testing();

