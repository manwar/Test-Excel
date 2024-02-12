#!/usr/bin/env perl

use strict;use warnings;
use Test::More;

eval "use Test::CheckManifest 0.9";
plan skip_all => "Test::CheckManifest 0.9 required" if $@;
ok_manifest({filter => [qr/\.git/]});
