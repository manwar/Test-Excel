package Test::Excel;

$Test::Excel::VERSION = '1.28';

=head1 NAME

Test::Excel - Interface to test and compare Excel files.

=head1 VERSION

Version 1.28

=cut

use strict; use warnings;

use 5.0006;
use IO::File;
use Data::Dumper;
use Test::Builder ();
use Scalar::Util 'blessed';
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::Utility qw(int2col col2int);

require Exporter;
our @ISA    = qw(Exporter);
our @EXPORT = qw(cmp_excel compare_excel);

$|=1;

my $DEBUG                = 0;
my $ALMOST_ZERO          = 10**-16;
my $IGNORE               = 1;
my $SPECIAL_CASE         = 2;
my $MAX_ERRORS_PER_SHEET = 0;

=head1 DESCRIPTION

This  module is meant to be used for testing  custom  generated  Excel  files, it
provides two functions at the moment, which is C<cmp_excel> and C<compare_excel>.
These can be used to compare_excel 2  Excel files to see  if they are I<visually>
similar. The function C<cmp_excel> is for testing purpose where function C<compare_excel>
can be used as standalone.

=head1 RULE

The new paramter has been added to both method cmp_excel() & method compare_excel()
called RULE. This is optional,however,this would allow to apply your own rule for
comparison. This should  be passed in as reference to a HASH with the keys sheet,
tolerance, sheet_tolerance and  optionally  swap_check,  error_limit  and message
(only relevant to method cmp_excel()).

    +-----------------+---------------------------------------------------------------------+
    | Key             | Description                                                         |
    +-----------------+---------------------------------------------------------------------+
    | sheet           | "|" seperated sheet names.                                          |
    | tolerance       | Number. Apply to all NUMBERS except on 'sheet'/'spec'. e.g. 10**-12 |
    | sheet_tolerance | Number. Apply to sheets/ranges in the spec. e.g. 0.20               |
    | spec            | Path to the specification file.                                     |
    | swap_check      | Number (optional) (1 or 0). Row swapping check. Default is 0.       |
    | error_limit     | Number (optional). Limit error per sheet. Default is 0.             |
    | message         | String (optional). Only required when calling method cmp_excel().   |
    +-----------------+---------------------------------------------------------------------+

=head1 SPECIFICATION FILE

Spec  file containing rules used should be in the format mentioned below. Key and
values are space seperated.

    sheet       Sheet1
    range       A3:B14
    range       B5:C5
    sheet       Sheet2
    range       A1:B2
    ignorerange B3:B8

=head1 What is "Visually" Similar?

This module uses the C<Spreadsheet::ParseExcel> module to parse Excel files, then
compares the parsed  data structure for differences. We ignore cetain  components
of the Excel file, such as embedded fonts,  images,  forms and  annotations,  and
focus  entirely  on  the layout of each Excel page instead.  Future versions will
likely support font and image comparisons, but not in this initial release.

=head1 METHODS

=head2 cmp_excel($got, $exp, { ...rule... })

This function will tell you whether the two Excel files are "visually" different,
ignoring  differences  in  embedded fonts/images and metadata. Both $got and $exp
can be either instances of Spreadsheet::ParseExcel / file path (which is in  turn
passed to the Spreadsheet::ParseExcel constructor). This one  is for use in  TEST
MODE.

    use strict; use warnings;
    use Test::More no_plan => 1;
    use Test::Excel;

    cmp_excel('foo.xls', 'bar.xls', { message => 'EXCELSs are identical.' });

    # or

    my $foo = Spreadsheet::ParseExcel::Workbook->Parse('foo.xls');
    my $bar = Spreadsheet::ParseExcel::Workbook->Parse('bar.xls');
    cmp_excel($foo, $bar, { message => 'EXCELs are identical.' });

=cut

sub cmp_excel {
    my ($got, $exp, $rule) = @_;

    _validate_rule($rule);
    $rule->{test} = 1;
    compare_excel($got, $exp, $rule);
}

=head2 compare_excel($got, $exp, { ...rule... })

This function will tell you whether the two Excel files are "visually" different,
ignoring differences in  embedded fonts/images and metadata. Both  $got and  $exp
can be either instances of Spreadsheet::ParseExcel / file path (which  is in turn
passed   to   the   Spreadsheet::ParseExcel constructor).  This one is for use in
STANDALONE MODE.

    use strict; use warnings;
    use Test::Excel;

    print "EXCELs are identical.\n"
    if compare_excel("foo.xls", "bar.xls");

=cut

sub compare_excel {
    my ($got, $exp, $rule) = @_;

    die("ERROR: Unable to locate file [$got].\n") unless (-f $got);
    die("ERROR: Unable to locate file [$exp].\n") unless (-f $exp);
    _log_message("INFO: Excel comparison [$got] [$exp]\n") if $DEBUG;

    unless (blessed($got) && $got->isa('Spreadsheet::ParseExcel::WorkBook')) {
        $got = Spreadsheet::ParseExcel::Workbook->Parse($got)
            || die("ERROR: Couldn't create Spreadsheet::ParseExcel::WorkBook instance with: [$got]\n");
    }

    unless (blessed($exp) && $exp->isa('Spreadsheet::ParseExcel::WorkBook')) {
        $exp = Spreadsheet::ParseExcel::Workbook->Parse($exp)
            || die("ERROR: Couldn't create Spreadsheet::ParseExcel::WorkBook instance with: [$exp]\n");
    }

    my (@gotWorkSheets, @expWorkSheets);
    my ($message, $status, $error, $error_limit, $spec, $test, $TESTER);

    $status = 1;
    $test = $rule->{test}                if ((ref($rule) eq 'HASH') && exists($rule->{test}));
    _validate_rule($rule)                unless (defined($test) && ($test));
    $spec = parse($rule->{spec})         if exists($rule->{spec});
    $error_limit = $rule->{error_limit}  if exists($rule->{error_limit});
    $message     = $rule->{message}      if exists($rule->{message});
    $error_limit = $MAX_ERRORS_PER_SHEET unless defined $error_limit;

    @gotWorkSheets = $got->worksheets();
    @expWorkSheets = $exp->worksheets();

    $TESTER = Test::Builder->new if (defined($test) && ($test));
    if (scalar(@gotWorkSheets) != scalar(@expWorkSheets)) {
        $error = "ERROR: Sheets count mismatch. ";
        $error .= "Got: [".scalar(@gotWorkSheets)."] exp: [".scalar(@expWorkSheets)."]\n";
        _log_message($error);
        if (defined($test) && ($test)) {
            $TESTER->ok(0, $message);
            return;
        }
        return 0;
    }

    my ($i, @sheets);
    @sheets = split(/\|/,$rule->{sheet})
        if (exists($rule->{sheet}) && defined($rule->{sheet}));

    for ($i=0; $i<scalar(@gotWorkSheets); $i++) {
        my ($error_on_sheet);
        my ($gotWorkSheet, $expWorkSheet);
        my ($gotSheetName, $expSheetName);
        my ($gotRowMin, $gotRowMax, $gotColMin, $gotColMax);
        my ($expRowMin, $expRowMax, $expColMin, $expColMax);

        $error_on_sheet = 0;
        $gotWorkSheet   = $gotWorkSheets[$i];
        $expWorkSheet   = $expWorkSheets[$i];
        $gotSheetName   = $gotWorkSheet->get_name();
        $expSheetName   = $expWorkSheet->get_name();
        if (uc($gotSheetName) ne uc($expSheetName)) {
            $error = "ERROR: Sheetname mismatch. Got: [$gotSheetName] exp: [$expSheetName].\n";
            _log_message($error);
            if (defined($test) && ($test)) {
                $TESTER->ok(0, $message);
                return;
            }
            return 0;
        }

        ($gotRowMin, $gotRowMax) = $gotWorkSheet->row_range();
        ($gotColMin, $gotColMax) = $gotWorkSheet->col_range();
        ($expRowMin, $expRowMax) = $expWorkSheet->row_range();
        ($expColMin, $expColMax) = $expWorkSheet->col_range();

        if ($DEBUG > 1) {
            _log_message("\n");
            _log_message("INFO:[$gotSheetName]:[$gotRowMin][$gotColMin]:[$gotRowMax][$gotColMax]");
            _log_message("INFO:[$expSheetName]:[$expRowMin][$expColMin]:[$expRowMax][$expColMax]");
        }

        if (defined($gotRowMax) && defined($expRowMax) && ($gotRowMax != $expRowMax)) {
            $error  = "\nERROR: Max row counts mismatch in sheet [$gotSheetName]. ";
            $error .= "Got[$gotRowMax] Expected: [$expRowMax]\n";
            _log_message($error);
            if (defined($test) && ($test)) {
                $TESTER->ok(0, $message);
                return;
            }
            return 0;
        }

        if (defined($gotColMax) &&  defined($expColMax) && ($gotColMax != $expColMax)) {
            $error  = "\nERROR: Max column counts mismatch in sheet [$gotSheetName]. ";
            $error .= "Got[$gotColMax] Expected: [$expColMax]\n";
            _log_message($error);
            if (defined($test) && ($test)) {
                $TESTER->ok(0, $message);
                return;
            }
            return 0;
        }

        my ($row, $col, $swap);
        for ($row=$gotRowMin; $row<=$gotRowMax; $row++) {
            for ($col=$gotColMin; $col<=$gotColMax; $col++) {
                my ($gotData, $expData, $error);
                $gotData = $gotWorkSheet->{Cells}[$row][$col]->{Val};
                $expData = $expWorkSheet->{Cells}[$row][$col]->{Val};

                next if ( defined($spec)
                          && exists($spec->{uc($gotSheetName)}->{$col+1}->{$row+1})
                          && ($spec->{uc($gotSheetName)}->{$col+1}->{$row+1} == $IGNORE) );

                if (defined($gotData) && defined($expData)) {
                    if (($gotData =~ /^[-+]?[0-9]*\.?[0-9]+([eE][-+]?[0-9]+)?$/)
                        && ($expData =~ /^[-+]?[0-9]*\.?[0-9]+([eE][-+]?[0-9]+)?$/)) {
                        if (($gotData < $ALMOST_ZERO) && ($expData < $ALMOST_ZERO)) {
                            # Can be treated as the same.
                            next;
                        }
                        else {
                            if (defined($rule)) {
                                my ($compare_with, $difference);
                                $difference = abs($expData - $gotData) / abs($expData);

                                if ( ( defined($spec)
                                       && exists($spec->{uc($gotSheetName)}->{$col+1}->{$row+1})
                                       && ($spec->{uc($gotSheetName)}->{$col+1}->{$row+1} == $SPECIAL_CASE)
                                     )
                                     ||
                                     ( scalar(@sheets) && grep(/$gotSheetName/,@sheets) ) ) {
                                    print "\nINFO: [NUMBER]:[$gotSheetName]:[SPC][".($row+1)."][".($col+1)."]:[$gotData][$expData] ... "
                                        if $DEBUG > 1;
                                    $compare_with = $rule->{sheet_tolerance};
                                }
                                else {
                                    print "\nINFO: [NUMBER]:[$gotSheetName]:[STD][".($row+1)."][".($col+1)."]:[$gotData][$expData] ... "
                                        if $DEBUG > 1;
                                    $compare_with = $rule->{tolerance};
                                }

                                if ($compare_with < $difference) {
                                    print "[FAIL]" if $DEBUG > 1;
                                    $difference = sprintf("%02f", $difference);
                                    $status = 0;
                                }
                                else {
                                    $status = 1;
                                    print "[PASS]" if $DEBUG > 1;
                                }
                            }
                            else {
                                print "\nINFO: [NUMBER]:[$gotSheetName]:[N/A][".($row+1)."][".($col+1)."]:[$gotData][$expData] ... "
                                    if $DEBUG > 1;
                                if ($expData != $gotData) {
                                    print "[FAIL]" if $DEBUG > 1;
                                    $status = 0;
                                }
                                else {
                                    $status = 1;
                                    print "[PASS]" if $DEBUG > 1;
                                }
                            }
                        }
                    }
                    else {
                        if (uc($gotData) ne uc($expData)) {
                            _log_message("INFO: [STRING]:[$gotSheetName]:[$expData][$gotData] ... [FAIL]");
                            $status = 0;
                        }
                        else {
                            $status = 1;
                            _log_message("INFO: [STRING]:[$gotSheetName]:[STD][".($row+1)."][".($col+1)."]:[$gotData][$expData] ... [PASS]")
                                if $DEBUG > 1;
                        }
                    }

                    if (exists($rule->{swap_check}) && defined($rule->{swap_check}) && ($rule->{swap_check})) {
                        if ($status == 0) {
                            $error_on_sheet++;
                            push @{$swap->{exp}->{number_to_letter($col-1)}}, $expData;
                            push @{$swap->{got}->{number_to_letter($col-1)}}, $gotData;

                            if (($error_on_sheet >= $error_limit) && ($error_on_sheet % 2 == 0) && !_is_swapping($swap)) {
                                _log_message("ERROR: Max error per sheet reached.[$error_on_sheet]\n");
                                if (defined($test) && ($test)) {
                                    $TESTER->ok($status, $message);
                                    return;
                                }
                                return $status;
                            }
                        }
                    }
                }
            } # col

            if (($error_on_sheet >= $error_limit) && ($error_on_sheet % 2 == 0) && !_is_swapping($swap)) {
                if (defined($test) && ($test)) {
                    $TESTER->ok($status, $message);
                    return;
                }
                return $status;
            }

        } # row

        if (exists($rule->{swap_check}) && defined($rule->{swap_check}) && ($rule->{swap_check})) {
            if (($error_on_sheet > 0) && _is_swapping($swap)) {
                print "\n\nWARN: SWAP OCCURRED.\n\n";
                $status = 1;
            }
        }
        print "INFO: [$gotSheetName]: ..... [OK].\n" if $DEBUG == 1;
    } # sheet


    if (defined($test) && ($test)) {
        $TESTER->ok($status, $message);
        return;
    }

    return $status;
}

sub parse {
    my ($spec) = @_;

    return unless defined $spec;

    die("ERROR: Unable to locate spec file [$spec].\n")
        unless (-f $spec);

    my ($handle, $row, $sheet, $cells, $data);
    $handle = IO::File->new($spec)
        || croak("ERROR: Couldn't open file [$spec][$!].\n");

    $sheet = undef;
    $data  = undef;
    while ($row = <$handle>) {
        chomp($row);
        next unless $row =~ /\w/;
        next if $row =~ /^#/;

        if ($row =~ /^sheet\s+(.*)/i) {
            $sheet = $1;
        }
        elsif (defined($sheet) && ($row =~ /^range\s+(.*)/i)) {
            $cells = Test::Excel::cells_within_range($1);
            foreach (@{$cells}) {
                $data->{uc($sheet)}->{$_->{col}+1}->{$_->{row}} = $SPECIAL_CASE;
            }
        }
        elsif (defined($sheet) && ($row =~ /^ignorerange\s+(.*)/i)) {
            $cells = Test::Excel::cells_within_range($1);
            foreach (@{$cells}) {
                $data->{uc($sheet)}->{$_->{col}+1}->{$_->{row}} = $IGNORE;
            }
        }
        else {
            die("ERROR: Invalid format data [$row] found in spec file.\n");
        }
    }
    $handle->close();

    return $data;
}

sub column_row {
    my ($cell) = @_;

    return unless defined $cell;

    die("ERROR: Invalid cell address [$cell].\n")
        unless ($cell =~ /([A-Za-z]+)(\d+)/);

    return ($1, $2);
}

sub letter_to_number {
    my ($letter) = @_;

    return col2int($letter);
}

sub number_to_letter {
    my ($number) = @_;

    return int2col($number);
}

sub cells_within_range {
    my ($range) = @_;

    return unless defined $range;

    die("ERROR: Invalid range [$range].\n")
        unless ($range =~ /(\w+\d+):(\w+\d+)/);

    my ($from, $to, $row, $col, $cells);
    my ($min_row, $min_col, $max_row, $max_col);

    $from = $1; $to = $2;
    ($min_col, $min_row) = column_row($from);
    ($max_col, $max_row) = column_row($to);
    $min_col = letter_to_number($min_col);
    $max_col = letter_to_number($max_col);

    for ($row = $min_row; $row <= $max_row; $row++) {
        for ($col = $min_col; $col <= $max_col; $col++) {
            push @{$cells}, { col => $col, row => $row };
        }
    }

    return $cells;
}

sub _is_swapping {
    my ($data) = @_;

    return 0 unless defined $data;

    foreach (keys %{$data->{exp}}) {
        my $exp = $data->{exp}->{$_};
        my $out = $data->{out}->{$_};

        return 0 if grep(/$exp->[0]/,@{$out});
    }

    return 1;
}

sub _log_message {
    my ($message) = @_;

    return unless defined($message);

    print {*STDOUT} "\n".$message;
}

sub _validate_rule {
    my ($rule) = @_;

    return unless defined $rule;

    die("ERROR: Invalid RULE definitions. It has to be reference to a HASH.\n")
        unless (ref($rule) eq 'HASH');

    my ($keys, $valid);
    $keys = scalar(keys(%{$rule}));
    return if (($keys == 1) && exists($rule->{message}));

    die("ERROR: Rule has more than 8 keys defined.\n")
        if $keys > 8;

    $valid = {'message'         => 1,
              'sheet'           => 2,
              'spec'            => 3,
              'tolerance'       => 4,
              'sheet_tolerance' => 5,
              'error_limit'     => 6,
              'swap_check'      => 7,
              'test'            => 8,};
    foreach (keys %{$rule}) {
        die("ERROR: Invalid key found in the rule definitions.\n")
            unless exists($valid->{$_});
    }

    if ((exists($rule->{spec}) && defined($rule->{spec}))
        || (exists($rule->{sheet}) && defined($rule->{sheet}))) {
        die("ERROR: Missing key sheet_tolerance in the rule definitions.\n")
            unless (exists($rule->{sheet_tolerance}) && defined($rule->{sheet_tolerance}));
        die("ERROR: Missing key tolerance in the rule definitions.\n")
            unless (exists($rule->{tolerance}) && defined($rule->{tolerance}));
    }
    else {
        if ( (exists($rule->{sheet_tolerance}) && defined($rule->{sheet_tolerance}))
             || (exists($rule->{tolerance}) && defined($rule->{tolerance})) ) {
            die("ERROR: Missing key sheet/spec in the rule definitions.\n")
                unless ((exists($rule->{sheet}) && defined($rule->{sheet}))
                        || (exists($rule->{spec}) && defined($rule->{spec})));
        }
    }
}

=head1 NOTES

It should be clearly noted that this module does not claim to provide  fool-proof
comparison of generated Excels. In fact there are still a number of ways in which
I want to expand the existing comparison functionality.This module I<is> actively
being developed for a number of projects  I  am  currently  working on, so expect
many  changes  to happen. If you have any suggestions/comments/questions   please
feel free to contact me.

=head1 CAVEATS

Testing of large Excels can take a long time, this is because, well, we are doing
a lot of computation. In fact, this   module   test  suite includes tests against
several  large  Excels,  however I am not including those in this distibution for
obvious reasons.

=head1 BUGS

None  that I am aware of. Of course, if you find a bug, let me know, and I will be
sure to fix it.  This is still a very early version, so it is always possible that
I have just "gotten it wrong" in some places.

=head1 SEE ALSO

=over 4

=item C<Spreadsheet::ParseExcel>  -  I  could  not have written this without this
module.

=back

=head1 ACKNOWLEDGEMENTS

=over 4

=item John McNamara (author of Spreadsheet::ParseExcel).

=item Kawai Takanori (author of Spreadsheet::ParseExcel::Utility).

=item Stevan Little (author of Test::PDF).

=back

=head1 AUTHOR

Mohammad S Anwar, C<< <mohammad.anwar at yahoo.com> >>

=head1 REPOSITORY

L<https://github.com/Manwar/Test-Excel>

=head1 BUGS

Please  report  any bugs or feature requests to C<bug-test-excel at rt.cpan.org>,
or through the web interface at L<http://rt.cpan.org/NoAuth/ReportBug.html?Queue=Test-Excel>.
I will be notified, and then you'll automatically be notified of progress on your
bug as I make changes.

=head1 SUPPORT

You can find documentation for this module with the perldoc command.

    perldoc Test::Excel

You can also look for information at:

=over 4

=item * RT: CPAN's request tracker (report bugs here)

L<http://rt.cpan.org/NoAuth/Bugs.html?Dist=Test-Excel>

=item * AnnoCPAN: Annotated CPAN documentation

L<http://annocpan.org/dist/Test-Excel>

=item * CPAN Ratings

L<http://cpanratings.perl.org/d/Test-Excel>

=item * Search CPAN

L<http://search.cpan.org/dist/Test-Excel/>

=back

=head1 LICENSE AND COPYRIGHT

Copyright (C) 2010 - 2014 Mohammad S Anwar.

This  program  is  free software; you can redistribute it and/or modify it under
the  terms  of the the Artistic License (2.0). You may obtain a copy of the full
license at:

L<http://www.perlfoundation.org/artistic_license_2_0>

Any  use,  modification, and distribution of the Standard or Modified Versions is
governed by this Artistic License.By using, modifying or distributing the Package,
you accept this license. Do not use, modify, or distribute the Package, if you do
not accept this license.

If your Modified Version has been derived from a Modified Version made by someone
other than you,you are nevertheless required to ensure that your Modified Version
 complies with the requirements of this license.

This  license  does  not grant you the right to use any trademark,  service mark,
tradename, or logo of the Copyright Holder.

This license includes the non-exclusive, worldwide, free-of-charge patent license
to make,  have made, use,  offer to sell, sell, import and otherwise transfer the
Package with respect to any patent claims licensable by the Copyright Holder that
are  necessarily  infringed  by  the  Package. If you institute patent litigation
(including  a  cross-claim  or  counterclaim) against any party alleging that the
Package constitutes direct or contributory patent infringement,then this Artistic
License to you shall terminate on the date that such litigation is filed.

Disclaimer  of  Warranty:  THE  PACKAGE  IS  PROVIDED BY THE COPYRIGHT HOLDER AND
CONTRIBUTORS  "AS IS'  AND WITHOUT ANY EXPRESS OR IMPLIED WARRANTIES. THE IMPLIED
WARRANTIES    OF   MERCHANTABILITY,   FITNESS   FOR   A   PARTICULAR  PURPOSE, OR
NON-INFRINGEMENT ARE DISCLAIMED TO THE EXTENT PERMITTED BY YOUR LOCAL LAW. UNLESS
REQUIRED BY LAW, NO COPYRIGHT HOLDER OR CONTRIBUTOR WILL BE LIABLE FOR ANY DIRECT,
INDIRECT, INCIDENTAL,  OR CONSEQUENTIAL DAMAGES ARISING IN ANY WAY OUT OF THE USE
OF THE PACKAGE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

=cut

1; # End of Test-Excel
