use strict; use warnings;
use Module::Build;

my $builder = Module::Build->new(
    sign           => 1,
    module_name    => 'Test::Excel',
    license        => 'perl',
    dist_author    => 'Mohammad S Anwar, <mohammad.anwar@yahoo.com>',
    add_to_cleanup => [ 'Test-Excel-*' ],	
    build_requires => {
        'Test::Builder'           => 0,
        'Spreadsheet::ParseExcel' => 0.58,
        'Scalar::Util'            => 0,
        'Test::Builder::Tester'   => 0,
        'Test::More'              => 0.47,
        'File::Spec'              => 0,
        'Readonly'                => 0,
    },
);

$builder->create_build_script();