#!/usr/bin/env perl

use strict;
use warnings;
use Modern::Perl;

use autodie qw{ :all };
use utf8;
use File::Spec;

use Smart::Comments '####';
use Getopt::Euclid;

use Text::CSV;
use File::Basename;
use Spreadsheet::ParseExcel;
use Text::Trim;

my $input_filename = $ARGV{'--infile'};
warn "Warning: The input file does not look like an Excel spreadsheet file."
    unless looks_like_xls($input_filename);
my $workbook = Spreadsheet::ParseExcel::Workbook->Parse($input_filename);

my ($inbase,$inpath,$insuffix) = fileparse($input_filename, qr{\.[^.]*});
my $output_base = File::Spec->catfile($inpath, $inbase);
my $csv = Text::CSV->new();

for my $worksheet ( $workbook->worksheets() ) { #### Processing worksheets (% done)...
    my ( $row_min, $row_max ) = $worksheet->row_range();
    my ( $col_min, $col_max ) = $worksheet->col_range();

    my @sheet_data;

    for my $row ( $row_min .. $row_max ) {
        for my $col ( $col_min .. $col_max ) {
            my $cell = $worksheet->get_cell( $row, $col ) or undef;
            next unless $cell;
            $sheet_data[$row][$col] = trim($cell->value());
        }
    }
    ### @sheet_data;

    my $ws_name = $worksheet->{Name};
    $ws_name =~ s{\s}{_}xsmg;   # no spaces
    my $csvname = "${output_base}-${ws_name}.csv";
    ### $csvname

    my $output = IO::File->new($csvname, '>') or die $!;
    ### $output

    foreach my $line (@sheet_data) {
        ### $line
        $csv->print($output, $line) or die $csv->error_diag();
        print $output "\n";
    }

    close $output;
}

sub extract_sheet_contents {
    my $sheet = $_[0];
    my ($nrow, $ncol) = ($sheet->{maxrow}, $sheet->{maxcol});
    return extract_rect_from_listref($sheet->{cell}, 1, $nrow, 1, $ncol);
}

sub extract_slice_of_listref {
    my ($listref, @slice) = @_;
    return [ map { $listref->[$_] } @slice ];
}

sub extract_rect_from_listref {
    my ($listref, $row_start, $row_end, $col_start, $col_end) = @_;
    return [ map {
        extract_slice_of_listref($_, $col_start..$col_end)
    } @{extract_slice_of_listref($listref, $row_start..$row_end)} ];
}

sub looks_like_xls {
    state $xls_regex = qr{\.xls$};
    return 1 if $_[0] =~ m{$xls_regex}i;
    return;
}


__END__

=head1 NAME

spreadsheet2csv-separate-sheets.pl - Split a spreadsheet into one csv file for each worksheet


=head1 VERSION

Version 1.0


=head1 USAGE

    progname [options]


=head1 REQUIRED ARGUMENTS

=over

=item --infile [=] <file> | -i <file>

The input spreadsheet file.

=for Euclid:
    file.type: readable
    file.default: '-'

=back


=head1 OPTIONS

=over

=item --version

=item --usage

=item --help

=item --man

Print the usual program information

=back

=head1 DESCRIPTION

This program will read a spreadsheet file and output one csv file for
each worksheet in the input file. The name of each output file will be
determined by the name of the input file and the name of the
worksheet. For example, a worksheet "Sheet1" in a file called
"reports.xls" will be output to "reports-Sheet1.csv".

=head1 NOTES

Empty rows and columns at the beginning of a worksheet will be
omitted. So if a worksheet has columns C through F filled, then the
output for that sheet will have exactly 4 columns, not 6.

=head1 AUTHOR

Ryan C. Thompson

=head1 BUGS

If you encounter a problem with this program, please email
rct+perlbug@thompsonclan.org. Bug reports and other feedback are
welcome.

=head1 COPYRIGHT AND LICENSE

This software is copyright (c) 2010 by Ryan C. Thompson.

This is free software; you can redistribute it and/or modify it under
the same terms as the Perl 5 programming language system itself.

