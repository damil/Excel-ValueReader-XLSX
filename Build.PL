use strict;
use warnings;
use Module::Build;

my $builder = Module::Build->new(
    module_name         => 'Excel::ValueReader::XLSX',
    license             => 'perl',
    dist_author         => 'Laurent Dami <dami@cpan.org>',
    dist_version_from   => 'lib/Excel/ValueReader/XLSX.pm',
    requires => {
      'perl'            => "5.010001",
      'Try::Tiny'       => 0,
     },
    build_requires => {
      'Test::More'   => 0,
    },
    add_to_cleanup      => [ 'Excel-ValueReader-XLSX-*' ],
    meta_merge => {
      resources => {
        repository => 'https://github.com/damil/Excel-ValueReader-XLSX',
       }
     },
);

$builder->create_build_script();
