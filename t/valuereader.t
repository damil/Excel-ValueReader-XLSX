use utf8;
use strict;
use warnings;
use Test::More;
use List::Util                qw/max/;
use Module::Load::Conditional qw/check_install/;

use Excel::ValueReader::XLSX;

(my $xl_file = $0) =~ s/\.t$/.xlsx/; # 'valuereader.xlsx' in the same directory

my @expected_sheet_names = qw/Test Empty Entities Tab_entities Dates/;
my @expected_values      = (  ["Hello", undef, undef, 22, 33, 55],
                              [123, undef, '<>'],
                              ["This is bold text", undef, '&'],
                              ["This is a Unicode string €", undef, '&<>'],
                              [],
                              [undef, "after an empty row and col",
                               undef, undef, undef,
                               "Hello after an empty row and col"],
                              ["cell\r\nwith\r\nembedded newlines"],
                             );

my @expected_tab_entities  = (
  [],
  [],
  ['Nombre de Name', "\x{c9}tiquettes de colonnes" ],
  ["\x{c9}tiquettes de lignes", 'capital', 'small', '(vide)',
   "Total g\x{e9}n\x{e9}ral"],
  ['A',                        '6',      '6',    undef, '12'],
  ['acute accent',             '1',      '1',    undef,  '2'],
  ['circumflex accent',        '1',      '1',    undef,  '2'],
  ['grave accent',             '1',      '1',    undef,  '2'],
  ['ring',                     '1',      '1',    undef,  '2'],
  ['tilde',                    '1',      '1',    undef,  '2'],
  ['dieresis or umlaut mark',  '1',      '1',    undef,  '2'],
  ['AE diphthong (ligature)',  '1',      '1',    undef,  '2'],
  ['(vide)',                   '1',      '1',    undef,  '2'],
  ['C',                        '1',      '1',    undef,  '2'],
  ['cedilla',                  '1',      '1',    undef,  '2'],
  ['E',                        '4',      '4',    undef,  '8'],
  ['acute accent',             '1',      '1',    undef,  '2'],
  ['circumflex accent',        '1',      '1',    undef,  '2'],
  ['grave accent',             '1',      '1',    undef,  '2'],
  ['dieresis or umlaut mark',  '1',      '1',    undef,  '2'],
  ['Eth',                      '1',      '1',    undef,  '2'],
  ['Icelandic',                '1',      '1',    undef,  '2'],
  ['greater than',             undef,    undef,    '1',  '1'],
  ['(vide)',                   undef,    undef,    '1',  '1'],
  ['I',                        '4',      '4',    undef,  '8'],
  ['acute accent',             '1',      '1',    undef,  '2'],
  ['circumflex accent',        '1',      '1',    undef,  '2'],
  ['grave accent',             '1',      '1',    undef,  '2'],
  ['dieresis or umlaut mark',  '1',      '1',    undef,  '2'],
  ['less than',                undef,    undef,    '1',  '1'],
  ['(vide)',                   undef,    undef,    '1',  '1'],
  ['N',                        '1',      '1',    undef,  '2'],
  ['tilde',                    '1',      '1',    undef,  '2'],
  ['O',                        '6',      '6',    undef, '12'],
  ['acute accent',             '1',      '1',    undef,  '2'],
  ['circumflex accent',        '1',      '1',    undef,  '2'],
  ['grave accent',             '1',      '1',    undef,  '2'],
  ['tilde',                    '1',      '1',    undef,  '2'],
  ['dieresis or umlaut mark',  '1',      '1',    undef,  '2'],
  ['slash',                    '1',      '1',    undef,  '2'],
  ['sharp s',                  undef,    '1',    undef,  '1'],
  ['German (sz ligature)',     undef,    '1',    undef,  '1'],
  ['single quote',             undef,    undef,    '1',  '1'],
  ['(vide)',                   undef,    undef,    '1',  '1'],
  ['THORN',                    '1',      '1',    undef,  '2'],
  ['Icelandic',                '1',      '1',    undef,  '2'],
  ['U',                        '4',      '4',    undef,  '8'],
  ['acute accent',             '1',      '1',    undef,  '2'],
  ['circumflex accent',        '1',      '1',    undef,  '2'],
  ['grave accent',             '1',      '1',    undef,  '2'],
  ['dieresis or umlaut mark',  '1',      '1',    undef,  '2'],
  ['Y',                        '1',      '2',    undef,  '3'],
  ['acute accent',             '1',      '1',    undef,  '2'],
  ['dieresis or umlaut mark',  undef,    '1',    undef,  '1'],
  ['(vide)',                   undef,    undef,    '1',  '1'],
  ['(vide)',                   undef,    undef,    '1',  '1'],
  ['ampersand',                undef,    undef,    '1',  '1'],
  ['(vide)',                   undef,    undef,    '1',  '1'],
  ["Total g\x{e9}n\x{e9}ral",  '30',     '32',     '5', '67'],
 );

my @expected_dates_and_times = (
  [ '10.07.2020',  '10.07.2020',  '01.02.1789', '10.07.2020 02:57:00', '02:57:59'],
  [ '10.07.2020',  '10.07.2020',  '31.12.1899', '10.07.2020 02:57:59', '01:23:00'],
  [ '10.07.2020',         undef,  '01.01.1900',                 undef, '01:26:18'],
  [ '10.07.2020',         undef,  '02.01.1900',                                  ],
  [ '10.07.2020',         undef,  '28.02.1900'                                   ],
  [ '10.07.2020',         undef,  '01.03.1900'                                   ],
  [ '10.07.2020',         undef,  '01.03.1900'                                   ],
  [ '10.07.2020',         undef,  '04.04.4444'                                   ],
  [ '10.07.2020'                                                                 ],
  [ '10.07.2020'                                                                 ],
  [ '10.07.2020'                                                                 ],
 );



my @backends = ('Regex');
push @backends, 'LibXML' if check_install(module => 'XML::LibXML::Reader');

foreach my $backend (@backends) {

  # dirty hack when testing with LibXML, because \r\n are silently transformed into \n
  local $expected_values[-1][0] = "cell\nwith\nembedded newlines"
    if $backend eq 'LibXML';

  # instantiate the reader
  my $reader = Excel::ValueReader::XLSX->new(xlsx => $xl_file, using => $backend);

  # check sheet names
  my @sheet_names = $reader->sheet_names;
  is_deeply(\@sheet_names, \@expected_sheet_names, "sheet names using $backend");

  # check a regular sheet
  my $values = $reader->values('Test');
  is_deeply($values, \@expected_values, "values using $backend");
  my $nb_cols = max map {scalar @$_} @$values;
  is ($nb_cols, 6, "nb_cols using $backend");

  # check an empty sheet
  my $empty  = $reader->values('Empty');
  is_deeply($empty, [], "empty values using $backend");

  # check a pivot table
  my $tab_entities = $reader->values('Tab_entities');
  is_deeply($tab_entities, \@expected_tab_entities, "tab_entities using $backend");

  # check date conversions
  my $dates = $reader->values('Dates');
  is_deeply($dates, \@expected_dates_and_times, "dates using $backend");
}

done_testing();

