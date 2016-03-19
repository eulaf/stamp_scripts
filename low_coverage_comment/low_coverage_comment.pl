#!/usr/bin/env perl

use strict;
use warnings;

use Tk;
use Tk::DropSite;
use Cwd qw(abs_path);
use File::Basename;
use Spreadsheet::Read;
use Carp;
#use Data::Dumper;

#-----------------------------------------------------------------------------

=head1 NAME

low_coverage_comment.pl - create low coverage comment

=head1 SYNOPSIS

B<low_coverage_comment> I<expression_file>

B<low_coverage_comment> I<depth_report1> I<depth_report2>

=head1 REQUIRES

Perl 5, 
Tk, 
Spreadsheet::Read, 
Spreadsheet::ParseExcel,
Spreadsheet::XLSX

=head1 OPTIONS

=over 4

=item B<-h>, B<--help>

Print short usage message.

=item B<--rst>

Print restructuredtext formatted output to stdout.

=back

=cut

#-----------------------------------------------------------------------------
### CONSTANTS

my $VERSION = "1.5";
my $BUILD = "160318";
my $OUTPUT_TEXT = "Portions of the following gene(s) failed to meet the".
" minimum coverage of MINCOVx: GENELIST. Low coverage may adversely affect".
" the sensitivity of the assay. If clinically indicated, repeat testing".
" on a new specimen can be considered.";
my @INPUT_CHOICES = qw/CSMP STAMP/;
my %MINCOVERAGE = ( CSMP=>300, STAMP=>200 );
my $MALE_MINCOV = 60;

#----------------

local $SIG{__WARN__} = sub {
    my $message = shift;
    croak "Error: " . $message;
};

my $DefaultText = "";
my $TextField = "";
my $TextField1 = "";
my $TextField2 = "";
my $UserInput;
my $UserInput2;
my $MainWindow;
my @Frames;
my $TextBox;
my $InputType;
my %Widgets = ();
my $RST = '';

if (@ARGV) { #check for commandline options
    if ($ARGV[0] =~ /^--?h(elp)*$/i) {
        print "Usage: $0 <expression report>\n";
        print "       $0 <depth report1> <depth report2>\n";
        print "version $VERSION\n";
        exit;
    } elsif ($ARGV[0] =~ /^--?r(st)*$/i) {
        $RST = ':';
        shift @ARGV;
    }
}

if (@ARGV) { 
    ($UserInput, $UserInput2)  = @ARGV;
    $InputType =  ($UserInput2) ? 'STAMP' : 'CSMP';
    run_program($InputType, $UserInput, $UserInput2);
} else { #run as GUI
    $MainWindow = MainWindow->new(-title=>'Low Coverage Comment Generator'.
                                          " (v$VERSION)");
    $MainWindow->geometry("620x570+0+0");
    $MainWindow->Label(-text=> "")->pack();
    my $radioframe = $MainWindow->Frame->pack;
    my $textframe = $MainWindow->Frame->pack(-ipadx=>10, -ipady=>10);
    my $csmpframe = $textframe->Frame->pack(-side=>'top');
    my $stampframe1 = $textframe->Frame->pack(-side=>'top');
    my $stampframe2 = $textframe->Frame->pack(-side=>'top');
    $Widgets{CSMP} = [$csmpframe];
    $Widgets{STAMP} = [$stampframe1, $stampframe2];
    setup_csmp_file_input($csmpframe, \$TextField);
    setup_stamp_file_input([$stampframe1, \$TextField1, \&accept_drop1,
            \&select_file_callback1], [$stampframe2, \$TextField2, 
            \&accept_drop2, \&select_file_callback2]);
    setup_radiobutton($radioframe);

    my $frame2 = $MainWindow->Frame;
    $frame2->Button(-width=>15, -text=>'Run', 
                    -command=>\&run_in_GUI)->pack(-side=>'left');
    $frame2->pack(-ipadx=>5, -ipady=>10);

    $TextBox = $MainWindow->Scrolled(qw' Text -height 25 -width 70', 
                                     -scrollbars=>'e')->pack;

    my $frame3 = $MainWindow->Frame;
    $frame3->Button( -width=>15, -text=>'Reset',
                     -command=>sub { 
                           undef $UserInput;
                           undef $UserInput2;
                           $TextField = $DefaultText;
                           $TextField1 = $DefaultText;
                           $TextField2 = $DefaultText;
                           $TextBox->delete("1.0",'end'); })
        ->pack(-side=>'left', -expand=>1);
    $frame3->Button(-width=>15, -text=>'Exit', 
                    -command=>sub { exit })->pack(-side=>'right', -expand=>1);
    $frame3->pack(-ipadx=>5, -ipady=>5);
    MainLoop();
} 

sub setup_radiobutton {
    my $radioframe = shift;
    foreach my $choice (@INPUT_CHOICES) {
        my $r = $radioframe->Radiobutton( -text=>$choice, 
                    -value=>$choice,
                    -variable=>\$InputType, 
                    -command=>[\&input_choice_made, $choice])
                 ->pack(-side=>'left');
        if ($choice eq $INPUT_CHOICES[0]) {
            $r->select();
            input_choice_made($choice)
        }
    }
}

sub input_choice_made {
    my $choice = shift;
    foreach my $k (keys %Widgets) {
        if ($k eq $choice) {
            pack_widgets($Widgets{$k});
        } else {
            forget_widgets($Widgets{$k});
        }
    }
}

sub pack_widgets {
    my $widgets = shift;
    foreach my $w (@$widgets) {
        $w->pack();
    }
}

sub forget_widgets {
    my $widgets = shift;
    foreach my $w (@$widgets) {
        $w->packForget();
    }
}

sub setup_csmp_file_input {
    my $frame = shift;
    my $textfieldref = shift;
    my $l = $frame->Label(-width=>15, -text=>"Expression report:");
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

sub setup_stamp_file_input {
    my @fieldlist = @_;
    my @widgetlist;
    foreach my $infields (@fieldlist) {
        my ($frame, $textfieldref, $drop, $select) = @$infields;
        my $l = $frame->Label(-width=>18, -text=>"STAMP depth report:");
        my $e = $frame->Entry(-width=>50, -textvariable=>$textfieldref);
        $e->DropSite(-dropcommand => [$drop, $e],
                -droptypes => ($^O eq 'MSWin32' ? 'Win32' : ['XDND', 'Sun']));
        my $b = $frame->Button(-width=>8, -text=>'Select file', 
                    -command=>[$select,]);
        $frame->pack();
        $l->pack(-side=>'left');
        $e->pack(-side=>'left');
        $b->pack(-side=>'left');
        push(@widgetlist, $l, $e, $b);
    }
    return;
}

#-----------------------------------------------------------------------------

sub run_program {
    my $inputtype = shift; #STAMP or CSMP
    my ($input, $input2) = @_;

    my $statuslist;
    my $fail = "FAIL";
    my $pass = "SUCCESS";
    my $mincoverage = $MINCOVERAGE{$inputtype};

    if (-f $input){ #a file
        if ($inputtype eq 'STAMP') {
            my @input = (abs_path($input),);
            if ($input2 && -f $input2) {
                push(@input, abs_path($input2));
            }
            eval {
                my $message = create_message_string($inputtype, \@input,
                      $MINCOVERAGE{$inputtype});
                foreach my $input (reverse sort @input) {
                    $statuslist->{$pass}->{$input} = $message;
                    $message = '';
                }

            };
        } else {
            eval {
                my $message = create_message_string($inputtype,
                        abs_path($input), $MINCOVERAGE{$inputtype});
                $statuslist->{$pass}->{$input} = $message;
            };
        }
        if ($@){
        print "FAIL $@.\n";
            $statuslist->{$fail}->{$input} = $@;
        }
    } else {
        print "No file to process.\n";
        $statuslist->{$fail}->{$input} = "No file to process.";
    }
    return $statuslist;
}

sub run_in_GUI {
    unless ($UserInput) {
        $TextBox->delete("1.0",'end');
        $TextBox->insert("1.0", "Please enter a valid expression report.");
        return;
    }
    my $statuslist;
    eval {
        my @inputs = $InputType eq 'CSMP' ? ($UserInput) : 
                                   ($UserInput, $UserInput2);
        $statuslist = run_program($InputType, @inputs);
    };
    if ($@) {
        $TextBox->delete("1.0",'end');
        $TextBox->insert("end", "Status: failed!\n");
        $TextBox->insert("end", "\n");
        $TextBox->insert("end", "$@");
    } else {
        $TextBox->delete("1.0",'end');
        if ($statuslist) { ##print FAIL message first
            for my $status (sort {$a cmp $b} keys %$statuslist){
                $TextBox->insert("end" , 
                                 "\n===============$status===============\n");
                for my $file (sort keys %{$statuslist->{$status}}){
                    $TextBox->insert("end", "Input file: $file\n");
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

sub create_message_string {
    my $inputtype = shift;
    my $inputfile = shift;
    my $mincoverage = shift;
    my $test = shift || 0;

    my ($low_genes, $numregions, $is_female);
    if ($inputtype eq 'STAMP') {
        ($low_genes, $numregions, $is_female) = get_low_cov_genes_stamp(
                                                $inputfile, $mincoverage);
    } elsif ($inputfile =~ /\.txt$/) {
        ($low_genes, $numregions) = get_low_cov_genes_txtfile($inputfile, 
                                                              $mincoverage);
    } else {
        ($low_genes, $numregions) = get_low_cov_genes($inputfile, $mincoverage);
    }
    my @low_genes = sort keys %$low_genes;
    my $outputtext = low_coverage_comment(\@low_genes, $mincoverage,
            $numregions);
    if ($is_female) {
        my @noYgenes = grep($$low_genes{$_} ne 'chrY', @low_genes);
        my $female_outtext = low_coverage_comment(\@noYgenes, $mincoverage);
        $outputtext = "All chrY regions have coverage < $MALE_MINCOV.\n".
                      "FEMALE:\n$female_outtext\n\nMALE:\n$outputtext";
    }
    if ($RST) {
        format_rst($outputtext);
    } else {
        print "\n$outputtext\n\n-----------------------------------\n\n";
    }
    return $outputtext;
}

sub low_coverage_comment {
    my $low_genes = shift;
    my $mincoverage = shift;
    my $numregions = shift;

    my $outputtext;
    if (@$low_genes) {
        my $genelist;
        if (@$low_genes <= 2) { 
            $genelist = join(" and ", @$low_genes)
        } else {
            my $lastgene = pop @$low_genes;
            $genelist = join(", ", @$low_genes);
            $genelist .= ", and $lastgene";
            push(@$low_genes, $lastgene);
        }
        $outputtext = $OUTPUT_TEXT;
        $outputtext =~ s/GENELIST/$genelist/;
        $outputtext =~ s/MINCOV/$mincoverage/;
    } else {
        $outputtext = "No low coverage genes in $numregions regions.";
    }
    return $outputtext;
}

sub get_low_cov_genes_stamp {
    my $inputfiles = shift;
    my $mincoverage = shift;

    my %genes;
    my $is_male = 0;
    my $numregions = 0;
    foreach my $inputfile (@$inputfiles) {
        print $RST."Input file: ".basename($inputfile)."\n";
        open(my $fh, "<", $inputfile) or croak "Could not read $inputfile: $!";
        my @lines = grep(/\w/, map { split(/[\n\r]/m, $_) } <$fh>);
        chomp @lines;
        $fh->close;
        my $mindepthfield;
        while (@lines) {
            if ($lines[0] =~ /Description.*\t(Min.Depth)\t/i) {
                $mindepthfield = $1;
                last;
            }
            shift @lines;
        }
        croak "Required columns 'Description', 'Min Depth' not found ".
            "in $inputfile\n" unless @lines;
        my @fields = split(/\t/, shift @lines);
        foreach my $l (@lines) {
            my @v = split(/\t/, $l);
            my %d = map { $_=> shift @v } @fields;
            if ($d{'Description'} && defined($d{$mindepthfield}) && 
                    $d{$mindepthfield} =~ /^\d+$/) {
                $numregions++;
                my $gene = $d{'Description'};
                $gene =~ s/_.*//;
                if ($d{$mindepthfield} < $mincoverage) {
                    $genes{$gene} = $d{Chr};
                }
                if ($d{Chr} eq 'chrY' and $d{$mindepthfield} >= $MALE_MINCOV) {
                    $is_male = 1;
                }
            } else {
                print "Bad $l\n";
            }
        }
    }
    return (\%genes, $numregions, !$is_male)
}

sub get_low_cov_genes_txtfile {
    my $inputfile = shift;
    my $mincoverage = shift;

    print $RST."Input file: ".basename($inputfile)."\n";
    open(my $fh, "<", $inputfile) or croak "Could not read $inputfile: $!";
    my @lines = grep(/\w/, map { split(/[\n\r]/m, $_) } <$fh>);
    chomp @lines;
    $fh->close;
    while (@lines) {
        last if $lines[0] =~ /Gene.*Min Counts/;
        shift @lines;
    }
    croak "Required columns 'Gene', 'Min Counts' not found in $inputfile\n" 
        unless @lines;
    my @fields = split(/\t/, shift @lines);
    my %genes;
    my $numregions = 0;
    foreach my $l (@lines) {
        my @v = split(/\t/, $l);
        my %d = map { $_=> shift @v } @fields;
        if ($d{'Gene'} && defined($d{'Min Counts'}) && 
                $d{'Min Counts'} =~ /^\d+$/) {
            $numregions++;
            my $gene = $d{'Gene'};
            $gene =~ s/;.*//;
            if ($d{'Min Counts'} < $mincoverage) {
                $genes{$gene} = $d{Chr};
            }
        } else {
            print "Bad $l\n";
        }
    }
    return (\%genes, $numregions)
}

sub get_low_cov_genes {
    my $inputfile = shift;
    my $mincoverage = shift;

    print $RST."Input file: ".basename($inputfile)."\n";
    my $wkbook = ReadData($inputfile) or
        croak "Failed to read $inputfile: $!\n";
    my $cells = $$wkbook[1]{cell}; # print Dumper($cells);
    my ($gene_col, $count_col, $chr_col);
    foreach my $col (@$cells) {
        if (grep(defined $_ && /Gene/, @$col)) {
            $gene_col = $col;
        } elsif (grep(defined $_ && /Min Counts/, @$col)) {
            $count_col = $col;
        } elsif (grep(defined $_ && /^Chr$/, @$col)) {
            $chr_col = $col;
        }
    }
    croak "Column 'Gene' not found in $inputfile\n" unless $gene_col;
    croak "Column 'Min Counts' not found in $inputfile\n" unless $count_col;
    my %genes;
    my $numregions = 0;
    for(my $i=0; $i<@$count_col; $i++) {
        next unless defined($$count_col[$i]) && 
              $$count_col[$i] =~ /^\d+$/;
        $numregions++;
        if ($$count_col[$i] < $mincoverage) {
            my $gene = $$gene_col[$i];
            $gene =~ s/;.*//;
            if ($chr_col) {
                $genes{$gene} = $$chr_col[$i];
            } else {
                $genes{$gene}++;
            }
        }
    }
    return (\%genes, $numregions)
}

sub format_rst {
    my $outputtext = shift;

    print "\n";
    print "+---------------------------------------+\n";
    printf "| %-37s |\n", 'v'.$VERSION;
    print "+=======================================+\n";
    my @lines = split(/\n/, $outputtext);
    foreach my $line (@lines) {
        if (length($line)>37) {
            my @words = split(/\s+/, $line);
            my $nextline = shift @words;
            my $reformattedline = '';
            foreach my $word (@words) {
                if (length($nextline) + length($word) < 37) {
                    $nextline .= " ".$word;
                } else {
                    $reformattedline .= sprintf "| %-37s |\n", $nextline;
                    $nextline = $word;
                }
            }
            $reformattedline .= sprintf "| %-37s |\n", $nextline;
            $reformattedline .= sprintf "| %-37s |\n", ' ';
            print $reformattedline;
        } else {
            unless ($line =~ /\w/) { $line = '|' }
            printf "| %-37s |\n", $line;
            printf "| %-37s |\n", '';
        }
    }
    print "+---------------------------------------+\n\n";
}

#-----------------------------------------------------------------------------

sub select_file_callback  {
    my $selected = $MainWindow->getOpenFile(
        -defaultextension => ".txt",
        -filetypes => [
                [ 'Text or Excel File', ['.txt', '.xls', '.xlsx']],
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

sub select_file_callback1  {
    my $selected = $MainWindow->getOpenFile(
        -defaultextension => ".txt",
        -filetypes => [
                [ 'Text only', ['.txt',]],
                ['All Files', '*']
            ],
        -title => "Open File"
    );
    if ($selected) {
        $UserInput = $selected;
        $TextField1 = $selected;
        $TextBox->delete("1.0",'end');
    }
}

sub select_file_callback2  {
    my $text_field = shift;
    my $selected = $MainWindow->getOpenFile(
        -defaultextension => ".txt",
        -filetypes => [
                [ 'Text only', ['.txt',]],
                ['All Files', '*']
            ],
        -title => "Open File"
    );
    if ($selected) {
        $UserInput2 = $selected;
        $TextField2 = $selected;
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

sub accept_drop1 {
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
        $TextField1 = $filename;
        $TextBox->delete("1.0",'end');
    }
}

sub accept_drop2 {
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
        $UserInput2 = $filename;
        $TextField2 = $filename;
        $TextBox->delete("1.0",'end');
    }
}

#-----------------------------------------------------------------------------

=head1 DESCRIPTION

Prints the appropriate low coverage comment for CSMP or STAMP runs.

For CSMP runs, the script takes a NextGENe expression report and 
prints the associated minimum coverage comment using a minimum coverage 
threshold of 300.

For STAMP runs, the scripts takes the STAMP depth reports for indels and
SNVs and prints the union of regions not meeting the minimum coverage
threshold of 200.

=head2 Input file

For CSMP, the input file must be an Excel file or tab-delimited text file 
in the format of an expression report generated by NextGENe.  The program 
looks for the 'Gene' and 'Min Counts' columns to determine which genes to 
include in the low coverage comment.

For STAMP, the input files should be tab-delimited text files.  The program
looks for the 'Description' and 'Min Depth' columns.  The 'Chr' column is
also used to determine if the sample is possibly female due to low 
coverage of chrY.

=head2 Output

The output is the low coverage comment string which will be printed to the
screen for copying and pasting elsewhere.  If all chrY genes have 
coverage < 60, two comments will be output: one for females excluding the chrY
genes and one for males, including all low coverage genes as usual.

=head1 REVISION HISTORY

=over 4

=item 1.4, 2015-07-29

Add check for female samples (all chrY regions having coverage < 60) and 
output additional comment excluding chrY genes if female.

=back

=cut
