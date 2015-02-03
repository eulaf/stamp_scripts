#!/bin/env perl
use strict;
use warnings;
#############################################################################
# File:     split_vcf.pl
# History:  Jan 2015    Rename and add ability to work with STAMP files 
#                       - Eula Fung
#           Oct 2014    Add documentation and testing - Eula Fung
#           Feb 2014    Rewrite by Xiaoqing You
#           5-Nov-2013  Original script by Chandler Ho
#############################################################################

=head1 NAME

split_vcf.pl - split VCF file into accepted and rejected

=head1 SYNOPSIS

B<split_vcf.pl> I<vcf_file> I<curated_mutation_file>

B<split_vcf.pl> I<vcf_file(s)> I<stamp_variant_file(s)>

=head1 DESCRIPTION

B<split_vcf.pl> takes a VCF file and curated mutation report or STAMP
variant report and creates two VCF files from the original VCF file:  one 
with accepted mutations and the other with the mutations that are not 
accepted.  Accepted mutations are determined from the mutation report with 
lines marked "vcf yes"; in the case of STAMP variant reports, all mutations
are accepted except those marked "NOT_REPORTED".

When running from the command line, multiple pairs of VCF files
and stamp variant reports can be run in batch, but they must have 
matching file basenames: 
e.g. unique_label.vcf, unique_label.variant_report.txt.

=cut

=head1 REQUIRES

Perl 5, 
Tk

=cut

use strict;
use Tk;
use Tk::DropSite;
use File::Basename;
use File::Spec;
use Cwd qw(abs_path);
use Class::Struct;
use IO::File;
use Carp;
use Data::Dumper;
use Getopt::Long;
use Pod::Usage;

##############################################################################
## CONSTANTS

my $PROGRAM = basename($0);
my $VERSION = "v1.0";
my $TITLE = "Split VCF $VERSION";

################################################################################
## MAIN PROGRAM

my $vtext =  "VCF file";
my $mtext = "mutation report or stamp variant report";
my $vcftext = $vtext;
my $mutationtext = $mtext;
my $mainwindow;
my $textbox;
my ($vcffile, $mutationfile);
my $p = { DBUG => 0, };


if (@ARGV==0) {
    setup_GUI();
} else {
    (my $infiles, $p) = setup_commandline();
    foreach my $label (sort keys %$infiles) {
        $vcffile = $$infiles{$label}{VCF};
        $mutationfile = $$infiles{$label}{REPORT};
        my $statuslist = run_program($vcffile, $mutationfile);
        if ($statuslist){ ##print FAIL message first
            for my $status (sort {$a cmp $b} keys %$statuslist){ 
                print STDERR "\n===============$status===============\n";
                for my $file (sort keys %{$statuslist->{$status}}){
                    print STDERR "$file\n";
                    print STDERR $statuslist->{$status}->{$file}. "\n\n";
                }
            }
        }
    }
}

################################################################################
## SUBROUTINES

sub setup_commandline {
    Pod::Usage::pod2usage(-verbose => 1) unless @ARGV;

    my $man = 0;
    my $help = 0;

    Getopt::Long::Configure('pass_through');
    Getopt::Long::GetOptions('help|?' => \$help, man => \$man);
    Pod::Usage::pod2usage(-verbose => 1,
                          -exitstatus => 0) if $help;
    Pod::Usage::pod2usage(-verbose => 2,
                          -exitstatus => 0) if $man;

    Getopt::Long::Configure('no_pass_through');
    Getopt::Long::GetOptions(
        'dbug' => \$$p{DBUG},
    ) || die "\n";
    @ARGV >= 2 || die "Need at least two arguments.\n";
    my %infiles;
    if (@ARGV==2) {
        my $vcf = shift @ARGV;
        my $report = shift @ARGV;
        $infiles{label}{VCF} = $vcf;
        $infiles{label}{REPORT} = $report;
    } else {
        my @VCF = grep(/\.vcf/i, @ARGV);
        my %reports = map { basename($_)=>$_ } grep(!/\.vcf/i, @ARGV);
        foreach my $vcf (@VCF) {
            my $label = basename($vcf, '.vcf');
            my @report = grep(/^$label\./, keys %reports);
            if (@report==1) {
                $infiles{$label}{VCF} = $vcf;
                $infiles{$label}{REPORT} = $reports{$report[0]};
            } else {
                die "Couldn't match $vcf to unique report @report\n";
            }
        }
    }
    (\%infiles, $p);
}


sub setup_GUI {
    $mainwindow = MainWindow->new(-title=>$TITLE);
    $mainwindow->geometry("800x625+0+0");
    $mainwindow->Label(-text=> "")->pack();
    my ($vcfentry, $vcfbutton, $mutationentry, $mutationbutton);

    my $frame1 = $mainwindow->Frame;
    my $frame1vcf = $frame1->Frame();
    $vcfbutton = $frame1vcf->Button(-width=>20, -text=>"VCF File",
                -state=>"normal",
                -command=>sub{ select_vcf_file_callback()})
        ->pack(-side=>'left', -expand=>1);
    $vcfentry = $frame1vcf->Entry(-width=>60, -textvariable=>\$vcftext,
                                  -state=>"normal" )
        ->pack(-side=>"right", -expand=>1);
    $vcfentry->DropSite(-dropcommand => [\&accept_drop, $vcfentry, \$vcftext,
                                         \$vcffile],
                        -droptypes => 
                            ($^O eq 'MSWin32' ? 'Win32' : ['XDND','Sun']));
    $frame1vcf->pack();

    my $frame1mutation = $frame1->Frame();
    $mutationbutton = $frame1mutation->Button(-width => 20,
            -text=>"Mutation Report File", -state=>"normal",
            -command=>sub{ select_mutation_file_callback(); })
            ->pack(-side=>'left', -expand=>1)
            ->pack(-side=>'left', -expand=>1);
    $mutationentry = $frame1mutation->Entry(-width => 60,
                                            -textvariable =>\$mutationtext,
                                            -state=>"normal" )
            ->pack(-side=>"right", -expand=>1);
    $mutationentry->DropSite(-dropcommand => [\&accept_drop, $mutationentry, 
                                              \$mutationtext, \$mutationfile],
                             -droptypes => ($^O eq 'MSWin32' ? 
                                       'Win32' : ['XDND','Sun']));
    $frame1mutation->pack();
    $frame1->pack(-ipadx => 10, -ipady => 10);

    my $frame2 = $mainwindow->Frame;
    $frame2->Button( -width=>15, -text =>'Run', -command => \&run_in_GUI)
        ->pack(-side=>'left', -expand=>1);
    $frame2->pack(-ipadx => 10, -ipady => 10);

    $textbox  =  $mainwindow->Scrolled(qw' Text -height 30 -width 95',
                                       -scrollbars=>'e')->pack;

    my $frame3 = $mainwindow->Frame;
    $frame3->Button( -width=>15, -text =>'Reset',
                     -command =>sub { $mutationtext=$mtext;
                         $vcftext =$vtext;
                         $textbox->delete("1.0", "end");
                         undef $vcffile;
                         undef $mutationfile; })
        ->pack(-side=>'left', -expand=>1);
    $frame3->Button(-width=>15,  -text =>'Exit',  -command => sub { exit })
        ->pack(-side=>'right', -expand=>1);
    $frame3->pack(-ipadx => 10, -ipady => 10);
    MainLoop();
}

sub run_program {
    my ($input1, $input2) = @_;
    my $statuslist;
    my $fail = "FAIL";
    my $pass = "SUCCESS";

    #input1 is VCF file, input2 is mutation report file
    if ($input2){
        eval {
            my ($afile, $rfile, $counts) = process_files($input1, $input2);
            $statuslist->{$pass}->{$afile} = 
                "$$counts{$afile} accepted variants";
            $statuslist->{$pass}->{$rfile} = 
                "$$counts{$rfile} rejected variants";
        };
        if ($@){
            $statuslist->{$fail}->{$input1} = $@;
        }
    }
    return $statuslist;
}

sub run_in_GUI {
    unless (($vcffile and -f $vcffile 
             and $mutationfile and -f $mutationfile) ) {
        $textbox->insert(0,
            "please enter a valid VCF and mutation report files");
        return;
    }
    $textbox->delete("1.0", "end");
    my $statuslist;
    eval {
        if ($vcffile and $mutationfile){
            $statuslist = run_program($vcffile, $mutationfile);
        }
    };
    if ($@){
        $textbox->insert("end", "Status: failed!\n\n");
        $textbox->insert("end",  "$@");
    }
    else {
        if ($statuslist){ ##print FAIL message first
            for my $status (sort {$a cmp $b} keys %$statuslist){ 
                $textbox->insert("end" , 
                        "\n===============$status===============\n");
                for my $file (sort keys %{$statuslist->{$status}}){
                    $textbox->insert("end", "$file\n");
                    $textbox->insert("end", $statuslist->{$status}->{$file}. 
                                     "\n\n");
                }
            }
        }
    }
}

sub parse_nextgene_mutation_file {
    my $fh = shift;

    my $colhash;
    my $poslist;
    my $hasheader = 0;
    while (<$fh>){
        next if (/^\s*$/);
        my @row = split(/\t/, $_);
        s/^\s+|\s+$//g for (@row);
        if (/^Index.+Chr/) {
            $hasheader=1;
            %$colhash = map {$row[$_]=>$_} 0..$#row;
            #print Data::Dumper->Dump([$colhash]); die;
        } elsif($hasheader) {
            if($_ =~/vcf\s+yes/i){
                my $chr = $row[$colhash->{'Chr'}];
                my $start = $row[$colhash->{'Chromosome Position'}];
                unless ($chr and $start){
                    next;
                }
                $chr =~ s/^\s+|\s+$//g;
                $start =~ s/^\s+|\s+$//g;
                my $genotype = $row[$colhash->{'Genotype'}];

                ##for INDELs (del or ins), use the previous position
                ## because VCF files describe indels using the previous pos
                ## while mutation files use the actual position
                if ($genotype =~ /^(del|ins)/i){
                    $start -= 1;
                }
                $poslist->{$chr}->{$start}++;
            }
        }
    }
    if (!$hasheader) {
        die "Badly formatted mutation report\n";
    }
    ($poslist)
}

sub parse_stamp_report {
    my $fh = shift;
    my $head = shift;

    my %poslist;
    $head =~ s/[\r\n]+$//;
    my @fields = split(/\t/, $head);
    while (my $line = <$fh>) {
        my @vals = split(/\t/, $line);
        $vals[$#vals] =~ s/[\n\r]+$//;
        my %d = map { $_=> shift @vals } @fields;
        my $info = "$d{Chr}\t$d{Position}\t$d{Gene}\t$d{Status}";
        if ($line =~ /NOT_REPORTED/ && $line !~ /EGFR/) {
            print "Rejected:\t$info\n";
            next;
        }
        next unless $d{Chr} && $d{Position};
        print Data::Dumper->Dump([\%d]) if $$p{DBUG};
        if ($d{Status} eq 'NOT_REPORTED' and $d{Gene} eq 'EGFR') {
            # Keep EGFR deletions
            if (length($d{Ref}) <= length($d{Var})) {
                print "Rejected:\t$info\n";
                next;
            }
        }
        print "Accepted:\t$info\n";
        my $chrom = $d{Chr};
        $chrom =~ s/chr//;
        $poslist{$chrom}{$d{Position}} = $line;
    }
    (\%poslist);
}

sub process_files {
    my ($vcffile, $mutationfile)  = @_;

    open(my $fh, "<", $mutationfile) or 
        die "Cannot open the file $mutationfile: $!\n";
    my $head = <$fh>;
    my $poslist = $head =~ /^Chr.*Position/ ? parse_stamp_report($fh, $head) :
                                    parse_nextgene_mutation_file($fh);
    $fh->close;
    my $basename = basename($vcffile);
    my $outdir = dirname($mutationfile);
    # create vcf accept and reject file
    my $acceptfile = $basename;
    $acceptfile  =~s/\.vcf//i;
    my $rejectfile  = $acceptfile;
    $acceptfile .='_accepted.vcf';
    $acceptfile = File::Spec->catfile($outdir, $acceptfile);
    $rejectfile .='_rejected.vcf';
    $rejectfile = File::Spec->catfile($outdir, $rejectfile);

    unless ($vcffile =~ /\.vcf$/i) {
        die "Expected .vcf extension for VCF file: $vcffile\n";
    }
    my $vcfio = IO::File->new("$vcffile") or
        croak "cannot open the file $vcffile: $!\n";
    my @vcflines = <$vcfio>;
    $vcfio->close();
    unless (grep(/^##fileformat=VCF/, @vcflines)) {
        die "Badly formatted VCF file: $vcffile.\n".
            "No ##fileformat line.\n";
    }
    my $acceptout = IO::File->new(">$acceptfile") or
        croak "cannot create the output file $acceptfile: $!\n";
    my $rejectout = IO::File->new(">$rejectfile") or
        croak "cannot create the output file $rejectfile: $!\n";

    my %counts = ($acceptfile=>0, $rejectfile=>0);
    foreach (@vcflines) {
        next if (/^\s*$/);
        if(/^\#/){#print header information into accept and reject files
            print $acceptout "$_";
            print $rejectout "$_";
        }
        else {
            my ($chr, $pos, undef) = split(/\t/,$_);
            $chr =~ s/^\s+|\s+$//g;
            $pos =~ s/^\s+|\s+$//g;
            if (exists $poslist->{$chr}->{$pos}){
                $counts{$acceptfile}++;
                print $acceptout "$_";
            } else {
                $counts{$rejectfile}++;
                print $rejectout "$_";
            }
        }
    }
    return ($acceptfile, $rejectfile, \%counts);
}


sub select_vcf_file_callback  {
    my $selected = $mainwindow->getOpenFile (
        -defaultextension => ".txt",
        -filetypes => [ [ 'VCF File', '.vcf'],
                        ['All Files', '*']
                       ],
        -title => "Open File"
    );
    if ($selected) {
        if (-f $selected){
            $vcftext = $selected;
            $vcffile = $selected;
            $textbox->delete("1.0",'end');
        } else {
            $textbox->delete("1.0",'end');
            $textbox->insert("1.0", "Please select a valid VCF file");
        }
    }
}

sub select_mutation_file_callback  {
    my $selected = $mainwindow->getOpenFile(
        -defaultextension => ".txt",
        -filetypes => [ [ 'Text File', '.txt'],
                        ['All Files', '*']
                      ],
        -title     => "Open File"
    );

    if ($selected) {
        if (-f $selected){
            $mutationtext = $selected;
            $mutationfile = $selected;
            $textbox->delete("1.0",'end');
        } else {
            $textbox->delete("1.0",'end');
            $textbox->insert("1.0", 
                    "please specify a valid mutation report file");
        }
    }
}


sub accept_drop {
    my $widget = shift;
    my $textfield = shift;
    my $filefield = shift;
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
        $$textfield = $filename;
        $$filefield = $filename;
        $textbox->delete("1.0",'end');
    }
}

