#!/usr/bin/perl

use strict; use warnings;

use Test::More;

use Test::Excel;
use File::Spec::Functions;

is(compare_excel(
    catfile('t', 'got-11.xls'),
    catfile('t', 'exp-11.xls'),
    { spec => catfile('t', 'spec-5.txt') }
  ), 1);

is(compare_excel(
    catfile('t', 'got-11.xls'),
    catfile('t', 'exp-11.xls'),
    { spec => catfile('t', 'spec-6.txt') }
  ), 1);

is(compare_excel(
    catfile('t', 'got-11.xls'),
    catfile('t', 'exp-11.xls'),
    { spec => catfile('t', 'spec-7.txt') }
  ), 0);

is(compare_excel(
    catfile('t', 'got-12.xls'),
    catfile('t', 'exp-12.xls'),
    { spec => catfile('t', 'spec-5.txt') }
  ), 0);

done_testing;
