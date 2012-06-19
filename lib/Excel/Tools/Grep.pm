#!/usr/bin/perl

use strict;
use warnings;
binmode(STDOUT, ":utf8");

use Spreadsheet::ParseExcel;
use Pod::Usage;
use Getopt::Long;

my %OPT = ();
GetOptions(\%OPT, 'man', 'i', 'file|f=s', 'help|h', 'verbose|v');
pod2usage("Please supply a pattern and a filename.") unless @ARGV;

my $parser = Spreadsheet::ParseExcel->new();

# get pattern or list of patterns
my @patterns;
if ($OPT{file}) {
    open(my $fh, '<', $OPT{file}) or die "$OPT{file}: $!";
    @patterns = map { chomp; $_ } grep { $_ =~ /\S/ } <$fh>;
}
else {
    push @patterns, shift @ARGV;
}

# workbooks
for my $file (@ARGV) {
    my $wb = $parser->parse($file);
    die($parser->error(), "\n") unless $wb;

    # get first worksheet only
    my $ws = $wb->worksheet(0);

    # get header
    my @header = get_header($ws);

    # look at all rows
    my ($rmin, $rmax) = $ws->row_range;
    my ($cmin, $cmax) = $ws->col_range;

    for my $row ($rmin .. $rmax) {
        my $question = do {
            my $cell = $ws->get_cell($row, 0);
            eval { $cell->value() } || q();
        };
        for my $col ($cmin .. $cmax) {
            my $cell = $ws->get_cell($row, $col);
            my $value = eval { $cell->value() } || q();
            my @m = grep {
                if ($OPT{i}) {
                    index(lc($value),lc($_)) >= 0
                }
                else {
                    index($value,$_) >= 0
                }
            } @patterns;
            next unless @m;
            show_match($file, \@header, $question, $row, $col, $value, \@m);
        }
    }
}

my $SHOWN;
sub show_match {
    my ($file, $header, $question, $row, $col, $value, $matches) = @_;
    my $field = $header->[$col];
    my $acol = alpha_col($col);
    $SHOWN = 1;
    print "=" x 40, "\n";
    local $, = q(,);
    print <<"__match"
file $file, question $question, field "$field"
matched strings: @$matches
cell value: $value
__match
}
END {
    if ($SHOWN) {
        print "=" x 40, "\n";
    }
    else {
        warn "No matches.\n";
    }
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

my $EXCEL_COLS;
sub alpha_col {
    my $col = shift;
    $EXCEL_COLS ||= ['A' .. 'Z', 'AA' .. 'ZZ'];
    return $EXCEL_COLS->[$col];
}

__END__

=pod

=head1 NAME

excel-grep - search for a string within Excel files

=head1 SYNOPSIS

    excel-grep --help
    excel-grep string file.xls [...]
    excel-grep --file pattern-file file.xls [...]

=head1 ARGUMENTS

=head1 -i

Performs string matching in a case-insensitive manner.

=cut

