#!/bin/env perl

# Copies variant reports of controls and renames them if necessary
# to include the STAMP run number.

use File::Basename;
use File::Copy;
use File::Spec;

use strict;

my $FORCE=0;
my $ARCHIVE = q"k:/0 STAMP PDF_reports_DataArchive/AnalysisFiles";
@ARGV || die "$0 <destdir>\n".
"\nCopies variant reports from\n$ARCHIVE\nto current directory.\n";

my $rundirs = get_rundirs($ARCHIVE);
my @rundirs = sort keys %$rundirs;
warn "Retrieved ".scalar(@rundirs)."\n";
my $i=0;
foreach my $rundir (@rundirs) { 
    my $report = find_control_variant_report($$rundirs{$rundir}, $rundir);
    my $reportname = $report ? basename($report) : 'Not found';
    printf STDERR "%5d %-s -- %-s\n", ++$i, $rundir, $reportname;
    next unless $report;
    my $runnum = $rundir; $runnum =~ s/STAMP(\d+)-.*$/$1/;
    my $newname = sprintf "TruQ3_%03d.variant_report.txt", $runnum;
#    if ($reportname ne $newname) { warn "      Rename as $newname\n"; }
    if (-f $newname && !$FORCE) {
        warn "      Already have $newname\n";
    } else {
        warn "      Copying $newname\n";
        copy($report, $newname);
    }
}

sub find_control_variant_report {
    my $dpath = shift;
    my $subdir = shift;

    my $reportsdir = File::Spec->catfile($dpath, 'reports');
    opendir(DIR, $reportsdir) || die "Dir $reportsdir: $!";
    my @files = grep(/^t.*q.*.variant_report.txt$/i, readdir(DIR));
    closedir(DIR);
    my $report;
    if (@files>1) {
        warn "Too many matches: ".join(", ", @files)."\n";
    } elsif (@files) {
        $report = File::Spec->catfile($reportsdir, shift @files);
    }
    return $report;
}

sub get_rundirs {
    my $dir = shift;
    opendir(DIR, $dir) or die "Could not read $dir: $!\n";
    my %seen;
    my %subdirs;
    foreach my $item (readdir(DIR)) {
        my $path = File::Spec->catfile($dir, $item);
        $path =~ /STAMP\d\d\d-analysis$/ or next;
        -d File::Spec->catfile($dir, $item, 'reports') or next;
        $subdirs{$item} = $path;
    }
    closedir(DIR);
    return \%subdirs;
}

__END__
my @stampdirs = grep(/STAMP/, get_subdirs('.'));
print "STAMP DIRS: ".join(", ", @stampdirs)."\n";

foreach my $subdir (@signoutdirs) {
    if (grep $subdir eq $_, @stampdirs) {
        warn "$subdir: already have\n";
    } else {
        my $dirpath = File::Spec->catfile($ARCHIVE, $subdir);
        my $pdfreports = get_pdf_reports($dirpath);
        if (%$pdfreports) {
            warn "$subdir:\n  ".join("\n  ", sort keys %$pdfreports)."\n";
            mkdir($subdir) or die "$subdir: $!";
            foreach my $pdf (sort keys %$pdfreports) {
                my $pdffile = $$pdfreports{$pdf};
                copy($pdffile, $subdir) or 
                    warn "Failed to copy $pdffile to $subdir: $!";
            }
        } else {
            warn "$subdir: No reports\n";
        }
    }
}

sub get_pdf_reports {
    my $dir = shift;

    opendir(DIR, $dir) or die "Could not read $dir: $!\n";
    my @subdirs = get_subdirs($dir);
    closedir(DIR);

    my %pdfreports;
    foreach my $subdir (@subdirs) {
        my $dirsubdir = File::Spec->catfile($dir, $subdir);
        opendir(DIR, $dirsubdir) || die "$dirsubdir: $!";
        my @pdf = grep(/\.pdf$/, readdir(DIR));
        closedir(DIR);
        foreach my $pdf (@pdf) {
            $pdfreports{$pdf} = File::Spec->catfile($dirsubdir, $pdf);
        }
    }
    return \%pdfreports;
}
