#!/usr/bin/perl

use strict;
use warnings;

use Spreadsheet::WriteExcel;
use File::Basename;

my $book = Spreadsheet::WriteExcel->new("C:\\Users\\Sundar\\Documents\\My\ Dropbox\\droidcon2011.xls")
                or die "Unable to create Excel file: $!";
my $sheet_name;
my $time_format = $book->add_format(num_format => '0.00', align => 'left');

for my $day_agenda (@ARGV) {
    unless($sheet_name = fileparse($day_agenda)) {
        print "$day_agenda is a directory, moving along";
        next;
    }
    my $sheet = $book->add_worksheet($sheet_name);
    open(my $fh, "<", $day_agenda) or die "Error reading $day_agenda, $!";
    my $prev_time = 0;
    my $row = 0;
    my $col = 0;
    while(<$fh>) {
        if (/^\s*(\d+):(\d+)(.*)/) {
            my $time = $1 . "." . $2;
            my $item = $3;
            if ($item =~ s/(\(.*?\))\s*$//) {
                $item = $1 . " " . $item;
            }
            if ($time == $prev_time) {
                $sheet->write(++$row, $col, $item);
            }
            else {
                $row = 0;
                $sheet->write($row, $col, $time, $time_format);
                $sheet->write(++$row, $col++, $item);
            }
            $prev_time = $time;
        }
    }
}
