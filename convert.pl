#!/usr/bin/env perl

use warnings;
use strict;

use Data::Dumper;

use File::Copy;
use File::Find::Rule;
use Spreadsheet::XLSX;
use Term::ANSIColor qw(:constants);

use constant SRC_IMG_LIST => 'images/list.txt';
use constant XLS_PRODUCTS => 'products/missing.xlsx';
use constant SRC_IMG_DIR  => '/Volumes/AFM_Artwork';
use constant TRG_IMG_DIR  => '/Volumes/Raid/MSC Drive/Product Images';

my $debug = 1;

mkdir(TRG_IMG_DIR) unless -d TRG_IMG_DIR;

if(!-f SRC_IMG_LIST || -z SRC_IMG_LIST) {
    _build_image_list(SRC_IMG_DIR);
}

open(FH, SRC_IMG_LIST); my @image_list = <FH>; close(FH);

if(scalar(@image_list) < 1) {
    print BOLD, RED, 'No images exist in: ', SRC_IMG_LIST, RESET, "\n";
    exit(1);
}

my $excel_parser = Spreadsheet::XLSX->new(XLS_PRODUCTS, my $converter);

foreach my $sheet(@{$excel_parser->{Worksheet}}) {
    foreach my $row(1 .. $sheet->{MaxRow}) {
        my $upc = $sheet->{Cells}[$row][0]->{Val};
        my $brand = $sheet->{Cells}[$row][2]->{Val};

        if($upc) {
            my $short_upc = substr($upc, -5);
            my $target_file = TRG_IMG_DIR.'/'.$upc.'.eps';
            my @upc_matches = grep(/${short_upc}\.eps$/, @image_list);

            chomp(@upc_matches);

            my $upc_count = scalar(@upc_matches);

            if($upc_count == 1) {
                if(!$debug) {
                    _copy_image($upc_matches[0], $target_file);
                    _convert_image($target_file);
                }
            } elsif($upc_count > 1) {
                if($brand) {
                    my @brand_matches = grep(/${brand}/, @upc_matches);

                    chomp(@brand_matches);

                    my $brand_count = scalar(@brand_matches);

                    if($brand_count == 1) {
                        if(!$debug) {
                            _copy_image($brand_matches[0], $target_file);
                            _convert_image($target_file);
                        }
                    } elsif($brand_count > 1) {
                        my @filenames;

                        foreach(@brand_matches) {
                            m/\/([^\/]+\.eps)/;
                            push(@filenames, $1);
                        }

                        if(scalar(_array_unique(@filenames)) == 1) {
                            if(!$debug) {
                                _copy_image($brand_matches[0], $target_file);
                                _convert_image($target_file);
                            }
                        } else {
                            my $product = $sheet->{Cells}[$row][6]->{Val};

                            if(!$product) {
                                $product = $brand;
                            }

                            print BOLD, WHITE, 'The following product has no definitive match:', RESET, "\n\n";
                            print '  ', BOLD, GREEN, 'UPC      ', RESET, $upc, "\n";
                            print '  ', BOLD, GREEN, 'BRAND    ', RESET, $brand, "\n";
                            print '  ', BOLD, GREEN, 'PRODUCT  ', RESET, $product, "\n\n";
                            print BOLD, WHITE, 'Please select the image you would like to use:', RESET, "\n\n";
                            print '  ', BOLD, WHITE, '0) ', 'Skip Product', RESET, "\n\n";

                            my $i = 1;

                            foreach(@brand_matches) {
                                m/\/([^\/]+\.eps)/;
                                print '  ', BOLD, BLUE, $i.') ', RESET, $1, "\n";
                                ++$i;
                            }

                            print "\n".': ';

                            is_numeric: {
                                my $number = <STDIN>;

                                chomp($number);

                                if($number !~ /^[0-9]+$/ || $number > $brand_count) {
                                    print BOLD, RED, "\n", 'Invalid selection. Please try again.', RESET, "\n\n", ': ';
                                    redo is_numeric;
                                } else {
                                    if($number > 0) {
                                        if(!$debug) {
                                            _copy_image($brand_matches[--$number], $target_file);
                                            _convert_image($target_file);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}

exit(0);

sub _array_unique {
    my %seen;

    grep(!$seen{$_}++, @_);
}

sub _copy_image {
    my($source, $target) = @_;

    print 'Copying ';
    print BOLD, WHITE, $source, RESET;
    print ' => ';
    print BOLD, BLUE, $target, RESET, "\n";

    copy($source, $target);

    return;
}

sub _convert_image {
    my $filename = shift;
    my $base_filename = substr($filename, 0, -4);

    print 'Converting image to JPG...';
    system('convert -density 300 -layers flatten "${filename}" "${base_filename}.jpg"');
    print BOLD, BLUE, 'done!', RESET, "\n";

    print 'Converting image to PNG...';
    system('convert -density 300 -layers flatten "${filename}" "${base_filename}.png"');
    print BOLD, BLUE, 'done!', RESET, "\n";
}

sub _build_image_list {
    my $dir = shift;

    print 'Building image list...';

    my @files = File::Find::Rule
        ->file()
        ->name(qr/_\d{5}\.eps$/)
        ->in($dir);

    open(FH, '>>', SRC_IMG_LIST);

    foreach my $file(@files) {
        if($file !~ /\/zz?\s?old\s?\/?/i) {
            say FH $file;
        }
    }

    close(FH);

    print BOLD, BLUE, 'done!', RESET, "\n";

    return;
}
