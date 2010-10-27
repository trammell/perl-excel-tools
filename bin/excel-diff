#!/usr/bin/perl

use strict;
use warnings;
binmode(STDOUT, ":utf8");

use Algorithm::Diff;
use String::Diff;
use Spreadsheet::ParseExcel;
use Pod::Usage;
use Getopt::Long;
use List::MoreUtils 'mesh';

my %OPT = ();
GetOptions(\%OPT, 'man', 'help|h', 'verbose|v');
pod2usage("Please supply two filenames.") unless @ARGV == 2;

my $parser = Spreadsheet::ParseExcel->new();

# workbooks
my $wb0= $parser->parse($ARGV[0]);
my $wb1= $parser->parse($ARGV[1]);

unless ($wb0 and $wb1) {
    die $parser->error(), "\n";
}

# worksheets
my $ws0 = $wb0->worksheet(0);
my $ws1 = $wb1->worksheet(0);

# diff headers
my @h0 = get_header($ws0);
my @h1 = get_header($ws1);

if ("@h0" ne "@h1") {
    warn "File headers differ.\n";
    my ($old, $new) = String::Diff::diff("@h0", "@h1");
    print "old: $old\n";
    print "new: $new\n";
}

# get question IDs (first column)
my %question_ids;
for my $ws ($ws0, $ws1) {
    my ($rmin, $rmax) = $ws->row_range;
    my ($cmin, $cmax) = $ws->col_range;
    #warn "rows: $rmin, $rmax";
    #warn "cols: $cmin, $cmax";

    for my $row ($rmin .. $rmax) {
        my $cell = $ws->get_cell($row, 0);
        my $value = eval { $cell->value() } || q();
        next if $value =~ /Question #/i;
        next unless $value =~ /\S/;
        $question_ids{ $value } = 1;
    }
}

# sort the question IDs correctly
my @question_ids = sort {
    my @a = ($a =~ /(\d+)-(\d+)/);
    my @b = ($b =~ /(\d+)-(\d+)/);
    $a[0] <=> $b[0] || $a[1] <=> $b[1];
} keys %question_ids;

# get differences
my @deltas;

for my $key (@question_ids) {
    warn "Diffing question ID '$key'\n" if $OPT{verbose};
    my $row0 = get_row($ws0,$key);
    my $row1 = get_row($ws1,$key);

    my @cols0 = @{$row0->{columns}};
    my @cols1 = @{$row1->{columns}};

    next if "@cols0" eq "@cols1";

    warn "Difference in ID $key\n" if $OPT{verbose};

    my %rec0 = mesh @h0, @cols0;
    my %rec1 = mesh @h1, @cols1;

    for my $h (@h0) {
        next if $rec0{$h} eq $rec1{$h};
        push @deltas, {
            key => $key,
            field => $h,
            old_row => $row0->{rownum},
            new_row => $row1->{rownum},
            old_val => $rec0{$h},
            new_val => $rec1{$h},
        };
    }
}

if (@deltas) {
    for my $delta (@deltas) {
        print "=" x 50, "\n";
        print "    file: ", $ARGV[1], "\n";
        print "question: ", $delta->{key}, "\n";
        print "     row: ", $delta->{new_row}, "\n";
        print "   field: ", $delta->{field}, "\n";
        print "     old: ", $delta->{old_val}, "\n";
        print "     new: ", $delta->{new_val}, "\n";
        my $n = first_diff($delta->{old_val}, $delta->{new_val});
        print "          ", " " x $n, "^\n";
    }
}

sub first_diff {
    my ($x,$y) = @_;
    my $xor = $x ^ $y;
    my ($nulls) = ($xor =~ /^([\000]*)/);
    return length($nulls);
}

sub get_row {
    my ($ws, $key) = @_;

    my ($rmin, $rmax) = $ws->row_range;
    my ($cmin, $cmax) = $ws->col_range;

    for my $row ($rmin .. $rmax) {
        my $cell = $ws->get_cell($row, 0);
        my $value = eval { $cell->value() } || q();
        next unless $value and $value eq $key;

        my @cols = map {
            my $cell = $ws->get_cell($row, $_);
            eval { $cell->value() } || q();
        } $cmin .. $cmax;
        return {
            rownum  => $row,
            columns => \@cols,
        };
    }
    warn "Key '$key' not found\n";
    return;
}

sub get_header {
    my $ws = shift;
    my ($cmin, $cmax) = $ws->col_range;
    my @header = map {
        my $cell = $ws->get_cell(0, $_);
        eval { $cell->value() } || q();
    } $cmin .. $cmax;
    return @header;
}

__END__

=pod

=head1 NAME

diff-excel - show the differences between two Excel files

=head1 SYNOPSIS

    diff-excel --help
    diff-excel file1.xls file2.xls

=head1 ARGUMENTS


=cut

# =head2 -k, --key-column
# diff-excel [--key=<column-name>] file1.xls file1.xls
# diff-excel [--key-column=<number>] file1.xls file1.xls
