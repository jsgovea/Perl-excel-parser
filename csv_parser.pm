use strict;
use warnings;
use Spreadsheet::Read;
use Excel::Writer::XLSX;

# Open the XLSX file
my $file          = "file_to_format.xls";
my $file_to_write = "FORMATO INV ISEP 2022.xlsx";

# File to read
my $workbook = Spreadsheet::Read->new($file);

# File to write
my $workbook_to_write  = ReadData($file_to_write);
my $worksheet_to_write = $workbook_to_write->[1];
my $workbook_styles    = $workbook_to_write->get_styles();

# Create a new workbook to write data while preserving styles
my $new_file      = "new_file.xlsx";
my $new_workbook  = Excel::Writer::XLSX->new($new_file);
my $new_worksheet = $new_workbook->add_worksheet();

# Check if the parsing was successful
die $workbook->error() if ( !defined $workbook );

# Read data from the first worksheet (index 1)
my $worksheet = $workbook->sheet(1);

# Read entire rows
for my $row ( 1 .. $worksheet->maxrow ) {
    my @row_data = $worksheet->row($row);

    my $style_properties = $workbook_styles->get_styles();
    my $new_style        = $new_workbook->add_format(%$style_properties);

    if ( $row_data[0] && $row_data[0] =~ /^(T8|T4)/ ) {
        my $inventory_number = $row_data[0];
        if ( $row > 17 ) {
            $new_worksheet->write( $row, 0, $inventory_number );
        }
    }
    if ( $row_data[3] ) {
        my $description = $row_data[3];
    }
    if ( $row_data[7] && $row < 207 ) {
        my $price = $row_data[7];
        print "$price\n";
    }
}

# $new_workbook->fo
$new_workbook->close();
