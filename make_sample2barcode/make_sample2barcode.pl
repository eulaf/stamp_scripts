#!/bin/env perl

=head1 NAME

make_sample2barcode.pl - create sample2barcode file from Excel spreadsheet

=head1 SYNOPSIS

B<make_sample2barcode> I<Excel sample key>

=head1 REQUIRES

Perl 5, 
File::Basename,
Tk, 
Tk::DropSite, 
Spreadsheet::Read, 
Spreadsheet::ParseExcel,
Spreadsheet::XLSX

=cut

use strict;
use warnings;

use Tk;
use Tk::DropSite;
use Cwd qw(abs_path);
use File::Basename;
use File::Spec;
use Spreadsheet::Read;
use Carp;
use Data::Dumper;

=head1 DESCRIPTION

Convert a Excel sample key to a STAMP sample2barcode.txt file.

=head2 Input file

The input file must be an Excel file with the name of the stamp run
in the format STAMP# (case-insensitive) somewhere in the file and
a table with columns: Name, lab# or acc#, mrn# and barcode (case-insensitive).
The control has no lab# or mrn#.

=head2 Output

The output is a tab-delimited sample2barcode_STAMP###.txt file for use 
with the STAMP analysis pipeline.  The output file will be written in
the same folder as the input file.

=head1 OPTIONS

=over 4

=item B<-h>, B<--help>

Print short usage message.

=back

=cut

### CONSTANTS

my $VERSION = "0.1";
my $BUILD = "150326";

#----------------

local $SIG{__WARN__} = sub {
    my $message = shift;
    croak "Error: " . $message;
 };

my $DefaultText = "";
my $TextField = "";
my $UserInput;
my $MainWindow;
my @Frames;
my $TextBox;
my $InputType;
my %Widgets = ();


sub setup_file_input {
    my $frame = shift;
    my $textfieldref = shift;
    my $l = $frame->Label(-width=>15, -text=>"STAMP sample key: ");
    my $e = $frame->Entry(-width=>50, -textvariable=>$textfieldref);
    $e->DropSite(-dropcommand => [\&accept_drop, $e ],
                -droptypes => ($^O eq 'MSWin32' ? 'Win32' : ['XDND', 'Sun']));
    my $b = $frame->Button(-width=>8, -text=>'Select file', 
                    -command=>[\&select_file_callback, $textfieldref]);
    $frame->pack();
    $l->pack(-side=>'left');
    $e->pack(-side=>'left');
    $b->pack(-side=>'left');
    return;
}

if (@ARGV) {
    ($UserInput)  = @ARGV;
    if ($UserInput =~ /^--?h(elp)*$/i) {
        usage_message();
    } else {
        print "Running program\n";
        run_program(@ARGV);
    }
}
elsif ($#ARGV < 0){#run as GUI
    $MainWindow = MainWindow->new(-title=>'STAMP sample2barcode generator'.
                                          " (v$VERSION)");
    $MainWindow->geometry("620x470+0+0");

    $MainWindow->Label(-text=> "")->pack();
    my $textframe = $MainWindow->Frame->pack(-ipadx=>10, -ipady=>10);
    my $fileframe = $textframe->Frame->pack(-side=>'top');
    setup_file_input($fileframe, \$TextField);

    my $frame2 = $MainWindow->Frame;
    $frame2->Button(-width=>15, -text=>'Run', 
                    -command=>\&run_in_GUI)->pack(-side=>'left');
    $frame2->pack(-ipadx=>5, -ipady=>10);

    $TextBox = $MainWindow->Scrolled(qw' Text -height 20 -width 70', 
                                     -scrollbars=>'e')->pack;

    my $frame3 = $MainWindow->Frame;
    $frame3->Button( -width=>15, -text=>'Reset',
                     -command=>sub { 
                           undef $UserInput;
                           $TextField = $DefaultText;
                           $TextBox->delete("1.0",'end'); })
        ->pack(-side=>'left', -expand=>1);
    $frame3->Button(-width=>15, -text=>'Exit', 
                    -command=>sub { exit })->pack(-side=>'right', -expand=>1);
    $frame3->pack(-ipadx=>5, -ipady=>5);
    MainLoop();
}
else { usage_message() }


sub usage_message {
    print "Usage: $0 <STAMP sample key>\n";
    print "version $VERSION\n";
    exit;
}

sub run_program {
    my @inputs = @_;

    my $statuslist;
    my $fail = "FAIL";
    my $pass = "SUCCESS";

    if (@inputs and -f $inputs[0]) {
        foreach my $input (@inputs) {
            eval {
                my $message = create_message_string(abs_path($input));
                $statuslist->{$pass}->{$input} = $message;
            };
            if ($@){
                print "FAIL $@.\n";
                $statuslist->{$fail}->{$input} = $@;
            }
        }
    } else {
        print "No file to process.\n";
        $statuslist->{$fail}->{$inputs[0]} = "No file to process.";
    }
    return $statuslist;
}

sub create_message_string {
    my $inputfile = shift;
    my $test = shift || 0;

    my ($runnum, $data) = get_sample_data($inputfile);
    my $contents = create_sample2barcode($runnum, $data);
    # If script is run from the K drive, then save data to
    # same directory as script; otherwise, save to same directory
    # as input file.
    my ($basename, $outpath) = fileparse($inputfile);
    if ($0 =~ /^K/) { ($basename, $outpath) = fileparse($0); }
    my $outfile = File::Spec->catfile($outpath, 
                  "sample2barcode_STAMP$runnum.txt");
    print STDERR "\nOutput file: $outfile\n";
    open(my $ofh, ">", $outfile) or die ">$outfile: $!";
    print $ofh $contents;
    close $ofh;
    return $contents . "\nContents written to $outfile\n";
}

sub get_sample_data {
    my $inputfile = shift;

    print "Reading $inputfile\n";
    unless ($inputfile =~ /\.xls/) {
        die "ERROR: Input $inputfile not an Excel file\n";
    }
    my $wkbook = ReadData($inputfile) or
        die "Failed to read $inputfile: $!\n";
    my @cells = grep @$_, @{ $$wkbook[1]{cell} }; 
    my %columns = (
            NAME => -1,
            LAB => -1,
            MRN => -1,
            BARCODE => -1);
    unless (@cells) {
        die "No data in wkbook\n";
    }
    my $runnum = '';
    my %data;
    for(my $i=0; $i<@cells; $i++) {
        my @columnvals = map { defined $_ ? $_ : '' } @{ $cells[$i] };
        # Find stamp run
        if (my @runnum = grep(/^stamp\s*\d+/i, @columnvals)) {
            $runnum[0] =~ s/stamp\s*(\d+)/$1/i;
            $runnum = sprintf "%03d", $runnum[0];
        }
        if (grep(/name/i, @columnvals)) {
            $data{name} = get_column_values('name', \@columnvals);
        } elsif (grep(/lab#|acc#/i, @columnvals)) {
            $data{lab} = get_column_values('lab#', \@columnvals);
        } elsif (grep(/mrn#/i, @columnvals)) {
            $data{mrn} = get_column_values('mrn#', \@columnvals);
        } elsif (grep(/barcode/i, @columnvals)) {
            $data{barcode} = get_column_values('barcode', \@columnvals);
        }
    }
    unless (exists $data{name}) { die "Name column not found\n"; }
    unless (exists $data{lab}) { die "Lab column not found\n"; }
    unless (exists $data{mrn}) { die "MRN column not found\n"; }
    unless (exists $data{barcode}) { die "Barcode column not found\n"; }
#    print Dumper(\@cells);
#    print Dumper([\%data]);
    return ($runnum, \%data);
}

sub get_column_values {
    my $field = shift;
    my $columnvals = shift;

    my $flag = 0;
    my %values;
    # Get values in column after field
    for(my $j=0; $j<@$columnvals; $j++) {
        if ($flag && $$columnvals[$j]) {
            $values{$j} = $$columnvals[$j];
            $values{$j} =~ s/^\s+|\s+$//;
        } elsif ($$columnvals[$j] =~ /$field/i) {
            $flag = 1;
        }
    }
    print STDERR "  $field: ".scalar(keys %values). " values\n";
    (\%values);
}

sub create_sample2barcode {
    my $runnum = shift;
    my $data = shift;

    my $contents = '';
    my @i = sort {$a<=>$b} keys %{ $$data{name} };

    foreach my $i (@i) {
        # Change name to last name + first initial(s)
        my $name = $$data{name}{$i};
        $name =~ s/,?[ _]([A-Z])[a-z]*/$1/g;
        $name =~ s/-//; #remove hyphens in hyphenated names
        print STDERR "$$data{name}{$i} --> $name\n";
        my $lab = $$data{lab}{$i} || '';
        my $mrn = $$data{mrn}{$i} || '';
        # control is special case w/o lab# or MRN
        my $sample = $lab || $mrn ?  join("_", $name, $lab, $mrn) : 
                     $name . '_'.$runnum;
        my $barcode = $$data{barcode}{$i};
        unless ($barcode) {
            print STDERR "No barcode for '$name'; SKIPPING\n";
            next;
        }
        $contents .= "$sample\t$barcode\n";
    }
    print STDERR "\n".$contents;
    ($contents);
}

#-----------------------------------------------------------------------------

sub run_in_GUI {
    unless ($UserInput) {
        $TextBox->delete("1.0",'end');
        $TextBox->insert("1.0", "Please enter a valid expression report.");
        return;
    }
    my $statuslist;
    eval {
        my @inputs = ($UserInput); 
        $statuslist = run_program(@inputs);
    };
    if ($@) {
        $TextBox->delete("1.0",'end');
        $TextBox->insert("end", "Status: failed!\n");
        $TextBox->insert("end", "\n");
        $TextBox->insert("end", "$@");
    }
    else {
        $TextBox->delete("1.0",'end');
        if ($statuslist) { ##print FAIL message first
            for my $status (sort {$a cmp $b} keys %$statuslist){
                $TextBox->insert("end" , 
                                 "\n===============$status===============\n");
                for my $file (sort keys %{$statuslist->{$status}}){
                    $TextBox->insert("end", "Input file: $file\n\n");
                    if ($statuslist->{$status}->{$file}) {
                        $TextBox->insert("end", "\n".
                                    $statuslist->{$status}->{$file}."\n\n\n");
                    }
                }
                $TextBox->insert("end" , "\n");
            }
        }
    }
}


sub select_file_callback  {
    my $selected = $MainWindow->getOpenFile(
        -defaultextension => ".xlsx",
        -filetypes => [
                [ 'Excel files', ['.xlsx','.xls']],
                ['All Files', '*']
            ],
        -title => "Open File"
    );
    if ($selected) {
        $UserInput = $selected;
        $TextField = $selected;
        $TextBox->delete("1.0",'end');
    }
}


sub accept_drop {
    my $widget = shift;
    my $selection = shift;

    my $filename;
    eval {
        if ($^O eq 'MSWin32') {
            $filename = $widget->SelectionGet(-selection => $selection,
                                              'STRING');
        } else {
            $filename = $widget->SelectionGet(-selection => $selection,
                                              'FILE_NAME');
        }
    };
    if (defined $filename) {
        $UserInput = $filename;
        $TextField = $filename;
        $TextBox->delete("1.0",'end');
    }
}

