#!/usr/bin/env perl

use warnings;
use strict;

use Data::Dumper;

use File::Copy;
use File::Find::Rule;
use Text::CSV_XS;
use Term::ANSIColor qw(:constants);

use constant SRC_CSV_FILE => 'data/afm.csv';
use constant TRG_TXT_FILE => 'data/afm.txt';
use constant TRG_MIS_FILE => 'data/missing.txt';
use constant SRC_IMG_DIR  => '/Volumes/AFM_Artwork';
use constant TRG_EPS_DIR  => '/Volumes/Raid/MSC Drive/Product Images';
use constant TRG_WEB_DIR  => '/Volumes/Raid/MSC Drive/Product Images/Web';

my $debug = 1;

if(!-f TRG_TXT_FILE || -z TRG_TXT_FILE) {
    buildImageList(SRC_IMG_DIR);
}

open(FH, TRG_TXT_FILE); my @imageList = <FH>; close(FH);

if(scalar(@imageList) < 1) {
    print BOLD, RED, 'No images exist in: ', TRG_TXT_FILE, RESET, "\n";
    exit(1);
}

my $csv = Text::CSV_XS->new({
    binary => 1,
    auto_diag => 1,
    allow_loose_quotes => 1,
    blank_is_undef => 1,
    sep_char => '|',
    escape_char => '\\',
    eol => $/
});

open(my $io, '<', SRC_CSV_FILE);

my @rows;

while(my $row = $csv->getline($io)) {
    push(@rows, $row);
}

$csv->eof or $csv->error_diag();

close($io);

for my $col(@rows) {
    my $upc = $col->[9];
    my $brand = $col->[3];

    if($upc) {
        my $shortUpc = $upc;
        my $epsImage = TRG_EPS_DIR.'/'.$upc.'.eps';

        if(length($upc) > 5) {
            $shortUpc = substr($upc, -5);
        }

        my @upcMatches = grep(/${shortUpc}\.eps$/, @imageList);

        chomp(@upcMatches);

        my $upcCount = scalar(@upcMatches);

        if($upcCount == 1) {
            if(!$debug) {
                copyImage($upcMatches[0], $epsImage);
                convertImage($epsImage);
            }
        } elsif($upcCount > 1) {
            if($brand) {
                my @brandMatches = grep(/${brand}/, @upcMatches);

                chomp(@brandMatches);

                my $brandCount = scalar(@brandMatches);

                if($brandCount == 1) {
                    if(!$debug) {
                        copyImage($brandMatches[0], $epsImage);
                        convertImage($epsImage);
                    }
                } elsif($brandCount > 1) {
                    my @filenames;

                    foreach(@brandMatches) {
                        m/\/([^\/]+\.eps)/;
                        push(@filenames, $1);
                    }

                    if(scalar(arrayUnique(@filenames)) == 1) {
                        if(!$debug) {
                            copyImage($brandMatches[0], $epsImage);
                            convertImage($epsImage);
                        }
                    } else {
                        my $product = $col->[2];

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
                                        copyImage($brandMatches[--$number], $epsImage);
                                        convertImage($epsImage);
                                    }
                                } else {
                                    addToMissingList($upc);
                                }
                            }
                        }
                    }
                } else {
                    addToMissingList($upc);
                }
            }
        } else {
            addToMissingList($upc);
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

    mkdir(TRG_EPS_DIR) unless -d TRG_EPS_DIR;

    print 'Copying ';
    print BOLD, WHITE, $source, RESET;
    print ' => ';
    print BOLD, BLUE, $target, RESET, "\n";

    copy($source, $target);

    return;
}

sub convertImage {
    shift =~ m/\/(\d+)\.eps$/;

    my $sourceImage = TRG_EPS_DIR.'/'.$1.'.eps';
    my $targetImage = TRG_WEB_DIR.'/'.$1.'.';

    mkdir(TRG_WEB_DIR) unless -d TRG_WEB_DIR;

    print 'Converting image to JPG...';
    system('convert -density 300 -layers flatten "'.$sourceImage.'" "'.$targetImage.'jpg"');
    print BOLD, BLUE, 'done!', RESET, "\n";

    print 'Converting image to PNG...';
    system('convert -density 300 -layers flatten "'.$sourceImage.'" "'.$targetImage.'png"');
    print BOLD, BLUE, 'done!', RESET, "\n";
}

sub addToMissingList {
    my $upc = shift;

    open(FH, '>>', TRG_MIS_FILE);
    say FH $upc;
    close(FH);
}

sub buildImageList {
    my $dir = shift;

    print 'Building image list...';

    my @files = File::Find::Rule
        ->file()
        ->name(qr/_\d{5}\.eps$/)
        ->in($dir);

    open(FH, '>>', TRG_TXT_FILE);

    foreach my $file(@files) {
        if($file !~ /\/zz?\s?old\s?\/?/i) {
            say FH $file;
        }
    }

    close(FH);

    print BOLD, BLUE, 'done!', RESET, "\n";

    return;
}
