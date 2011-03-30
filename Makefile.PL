use ExtUtils::MakeMaker;
WriteMakefile(
    NAME         => 'Test::Excel',
    VERSION_FROM => 'lib/Test/Excel.pm', 
    clean        => { FILES => 'Test-Excel-*' },
    PREREQ_PM    => {
    'Test::Builder'           => 0,
    'Spreadsheet::ParseExcel' => 0.58,
    'Scalar::Util'            => 0,
    'Test::Builder::Tester'   => 0,
    'Test::More'              => 0.47,
    'File::Spec'              => 0,
    'Readonly'                => 0,
    }
);