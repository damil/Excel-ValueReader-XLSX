use utf8;
use strict;
use warnings;
use Getopt::Long;

use Excel::ValueReader::XLSX;
# use Excel::Reader::XLSX;
use Spreadsheet::ParseXLSX;
use Data::XLSX::Parser;

# options de la ligne de commande
GetOptions \my %opt, 
  'xl_file=s',              # fichier Excel des comparaisons
  'valuereader!',
  'ivaluereader!',
  'vrlibxml!',
  'ivrlibxml!',
  'reader!',
  'parsexlsx!',
  'xparser!',
  ;

$opt{xl_file} //= "d:/temp/Audit/Stats_acces_2024/Stats_DM_par_semaine_S1_2024.xlsx";


my ($start, $cpu, $system) = (time, times);

valuereader($opt{xl_file}) if $opt{valuereader};
ivaluereader($opt{xl_file})if $opt{ivaluereader};
vrlibxml($opt{xl_file})    if $opt{vrlibxml};
ivrlibxml($opt{xl_file})   if $opt{ivrlibxml};
reader($opt{xl_file})      if $opt{reader};
parsexlsx($opt{xl_file})   if $opt{parsexlsx};
xparser($opt{xl_file})     if $opt{xparser};

my ($end, $ecpu, $esystem) = (time, times);

printf "%d elapsed, %d cpu, %d system\n", $end-$start, $ecpu-$cpu, $esystem-$system;

sub valuereader {
  my $xl_file = shift;

  warn "using ValueReader\n";
  my $rv = Excel::ValueReader::XLSX->new(xlsx => $xl_file);
  foreach my $sheet_name ($rv->sheet_names) {
    my $vals = $rv->values($sheet_name);
    my $n_rows = @$vals;
    warn "sheet $sheet_name has $n_rows rows\n";
  }
}


sub ivaluereader {
  my $xl_file = shift;

  warn "using ValueReader iterator\n";
  my $rv = Excel::ValueReader::XLSX->new(xlsx => $xl_file);
  foreach my $sheet_name ($rv->sheet_names) {
    my $it = $rv->ivalues($sheet_name);
    my $n_rows = 0;
    $n_rows++ while $it->();
    warn "sheet $sheet_name has $n_rows rows\n";
  }
}






sub vrlibxml {
  my $xl_file = shift;

  warn "using ValueReader with LibXML\n";
  my $rv = Excel::ValueReader::XLSX->new(xlsx => $xl_file, using => 'LibXML');
  foreach my $sheet_name ($rv->sheet_names) {
    my $vals = $rv->values($sheet_name);
    my $n_rows = @$vals;
    warn "sheet $sheet_name has $n_rows rows\n";
  }
}


sub ivrlibxml {
  my $xl_file = shift;

  warn "using ValueReader iterator with LibXML\n";
  my $rv = Excel::ValueReader::XLSX->new(xlsx => $xl_file, using => 'LibXML');
  foreach my $sheet_name ($rv->sheet_names) {
    my $it = $rv->ivalues($sheet_name);
    my $n_rows = 0;
    $n_rows++ while $it->();
    warn "sheet $sheet_name has $n_rows rows\n";
  }
}

sub reader {
  my $xl_file = shift;

  warn "using Excel::Reader::XLSX\n";
  my $reader   = Excel::Reader::XLSX->new();
  my $workbook = $reader->read_file($xl_file);
  for my $worksheet ( $workbook->worksheets() ) {
    my $sheet_name = $worksheet->name();
    my @rows;
    while ( my $row = $worksheet->next_row() ) {
      my @row;
      while ( my $cell = $row->next_cell() ) {
        push @row, $cell->value();
      }
      push @rows, \@row;
    }
    my $n_rows = @rows;
    warn "sheet $sheet_name has $n_rows rows\n";
  }
}


sub parsexlsx {
  my $xl_file = shift;

  warn "using Spreadsheet::ParseXLSX\n";
  my $parser    = Spreadsheet::ParseXLSX->new();
  my $workbook  = $parser->parse($xl_file) or die $parser->error;
  for my $worksheet ( $workbook->worksheets() ) {
    my $sheet_name = $worksheet->get_name();
    my ( $row_min, $row_max ) = $worksheet->row_range();
    warn "sheet $sheet_name has $row_max rows\n";
  }
}



sub xparser {
  my $xl_file = shift;

  warn "using Data::XLSX::Parser\n";

  my $parser = Data::XLSX::Parser->new;
  my @rows;
  $parser->add_row_event_handler(sub {
    my ($row) = @_;
    push @rows, $row;
  });

  $parser->open($xl_file);

  foreach my $sheet_name ($parser->workbook->names) {
    @rows = ();
    $parser->sheet_by_rid( "rId" . $parser->workbook->sheet_id( $sheet_name ) );
    my $n_rows = @rows;
    warn "sheet $sheet_name has $n_rows rows\n";
    @rows = ();
  }
}



__END__

using ValueReader
sheet Stats_DM_par_semaine_S1_2024 has 800131 rows
40 elapsed, 32 cpu, 0 system

using ValueReader iterator
sheet Stats_DM_par_semaine_S1_2024 has 800131 rows
34 elapsed, 30 cpu, 0 system

using ValueReader with LibXML
sheet Stats_DM_par_semaine_S1_2024 has 800131 rows
101 elapsed, 83 cpu, 0 system


using ValueReader iterator with LibXML
sheet Stats_DM_par_semaine_S1_2024 has 800131 rows
91 elapsed, 80 cpu, 0 system



using Spreadsheet::ParseXLSX
sheet Stats_DM_par_semaine_S1_2024 has 800130 rows
1272 elapsed, 870 cpu, 4 system

using Data::XLSX::Parser
sheet Stats_DM_par_semaine_S1_2024 has 800131 rows
125 elapsed, 107 cpu, 1 system
  
