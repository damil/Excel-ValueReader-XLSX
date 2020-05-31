use utf8;
use strict;
use warnings;
use Test::More;
use Module::Load::Conditional qw/check_install/;

use lib "../lib";
use Excel::ValueReader::XLSX;

(my $xl_file = $0) =~ s/\.t$/.xlsx/;

my @expected_sheet_names = qw/Test Empty/;
my @expected_values      = (  ["Hello", undef, undef, 22, 33, 55],
                              [123],
                              ["This is bold text"],
                              ["This is a Unicode string €"],
                              undef,
                              [undef, "after an empty row and col",
                               undef, undef, undef,
                               "Hello after an empty row and col"],
                             );
my @backends = ('Regex', check_install(module => 'XML::LibXML::Reader') ? 'Libxml' : ());

foreach my $backend (@backends) {

  my $reader = Excel::ValueReader::XLSX->new(xlsx => $xl_file, using => $backend);
  my @sheet_names = $reader->sheet_names;
  is_deeply(\@sheet_names, \@expected_sheet_names, "sheet names using $backend");

  my $values = $reader->values('Test');
  is_deeply($values, \@expected_values, "values using $backend");
}

done_testing();

