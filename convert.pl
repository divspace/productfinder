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

my $debug = 0;

mkdir(TRG_IMG_DIR) unless -d TRG_IMG_DIR;

if(!-f SRC_IMG_LIST || -z SRC_IMG_LIST) {
    buildImageList(SRC_IMG_DIR);
}

open(FH, SRC_IMG_LIST); my @imageList = <FH>; close(FH);

if(scalar(@imageList) < 1) {
    print BOLD, RED, 'No images exist in: ', SRC_IMG_LIST, RESET, "\n";
    exit(1);
}

my $excelParser = Spreadsheet::XLSX->new(XLS_PRODUCTS, my $converter);

foreach my $sheet(@{$excelParser->{Worksheet}}) {
    foreach my $row(1 .. $sheet->{MaxRow}) {
        my $upc = $sheet->{Cells}[$row][0]->{Val};
        my $brand = $sheet->{Cells}[$row][2]->{Val};

        if($upc) {
            my $shortUpc = substr($upc, -5);
            my $targetFile = TRG_IMG_DIR.'/'.$upc.'.eps';
            my @upcMatches = grep(/${shortUpc}\.eps$/, @imageList);

            chomp(@upcMatches);

            my $upcCount = scalar(@upcMatches);

            if($upcCount == 1) {
                if(!$debug) {
                    copyImage($upcMatches[0], $targetFile);
                    convertImage($targetFile);
                }
            } elsif($upcCount > 1) {
                if($brand) {
                    my @brandMatches = grep(/${brand}/, @upcMatches);

                    chomp(@brandMatches);

                    my $brandCount = scalar(@brandMatches);

                    if($brandCount == 1) {
                        if(!$debug) {
                            copyImage($brandMatches[0], $targetFile);
                            convertImage($targetFile);
                        }
                    } elsif($brandCount > 1) {
                        my @filenames;

                        foreach(@brandMatches) {
                            m/\/([^\/]+\.eps)/;
                            push(@filenames, $1);
                        }

                        if(scalar(arrayUnique(@filenames)) == 1) {
                            if(!$debug) {
                                copyImage($brandMatches[0], $targetFile);
                                convertImage($targetFile);
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

                            foreach(@brandMatches) {
                                m/\/([^\/]+\.eps)/;
                                print '  ', BOLD, BLUE, $i.') ', RESET, $1, "\n";
                                ++$i;
                            }

                            print "\n".': ';

                            isNumeric: {
                                my $number = <STDIN>;

                                chomp($number);

                                if($number !~ /^[0-9]+$/ || $number > $brandCount) {
                                    print BOLD, RED, "\n", 'Invalid selection. Please try again.', RESET, "\n\n", ': ';
                                    redo isNumeric;
                                } else {
                                    if($number > 0) {
                                        if(!$debug) {
                                            copyImage($brandMatches[--$number], $targetFile);
                                            convertImage($targetFile);
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

sub arrayUnique {
    my %seen;

    grep(!$seen{$_}++, @_);
}

sub copyImage {
    my($source, $target) = @_;

    print 'Copying ';
    print BOLD, WHITE, $source, RESET;
    print ' => ';
    print BOLD, BLUE, $target, RESET, "\n";

    copy($source, $target);

    return;
}

sub convertImage {
    my $filename = shift;
    my $baseFilename = substr($filename, 0, -4);

    print 'Converting image to JPG...';
    system('convert -density 300 -layers flatten "${filename}" "${baseFilename}.jpg"');
    print BOLD, BLUE, 'done!', RESET, "\n";

    print 'Converting image to PNG...';
    system('convert -density 300 -layers flatten "${filename}" "${baseFilename}.png"');
    print BOLD, BLUE, 'done!', RESET, "\n";
}

sub buildImageList {
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
