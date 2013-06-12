use strict;
use warnings;
use Module::Build;

my $builder = Module::Build->new(
    module_name         => 'Catmandu::Exporter::XLS',
    license             => 'perl',
    dist_author         => q{Nicolas Steenlant <nicolas.steenlant@ugent.be>},
    dist_version_from   => 'lib/Catmandu/Exporter/XLS.pm',
    release_status      => 'stable',
    configure_requires => {
        'Module::Build' => 0,
    },
    build_requires => {
        'Test::More' => 0,
    },
    requires => {
        'Catmandu' => '0.2003'
        'Spreadsheet::WriteExcel' => '2.37'
    },
    add_to_cleanup     => [ 'Catmandu-Exporter-XLS-*' ],
    create_makefile_pl => 'traditional',
);

$builder->create_build_script();
