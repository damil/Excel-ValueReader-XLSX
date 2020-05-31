package Excel::ValueReader::XLSX::Regex;
use utf8;
use Moose;

#======================================================================
# GLOBAL VARIABLES
#======================================================================

our $VERSION = '1.0';

my %xml_entities   = ( amp  => '&',
                       lt   => '<',
                       gt   => '>',
                       quot => '"',
                       apos => "'",  );
my $entity_names   = join '|', keys %xml_entities;
my $regex_entities = qr/&($entity_names);/;

#======================================================================
# ATTRIBUTES
#======================================================================

has 'frontend'  => (is => 'ro',   isa => 'Excel::ValueReader::XLSX', 
                    required => 1, weak_ref => 1,
                    handles => [qw/sheet_member _member_contents strings A1_to_num/]);

#======================================================================
# LAZY ATTRIBUTE CONSTRUCTORS
#======================================================================

sub _strings {
  my $self = shift;
  my @strings;

  # read from the sharedStrings zip member
  my $contents = $self->_member_contents('xl/sharedStrings.xml');

  # iterate on <si> nodes
  while ($contents =~ m[<si>(.*?)</si>]g) {
    my $innerXML = $1;

    # concatenate contents from all <t> nodes (usually there is only 1)
    my $string   = join "", ($innerXML =~ m[<t[^>]*>(.+?)</t>]g);

    # decode entities
    $string =~ s/$regex_entities/$xml_entities{$1}/eg;

    push @strings, $string;
  }

  return \@strings;
}



sub _sheets {
  my $self = shift;

  # read from the workbook.xml zip member
  my $contents = $self->_member_contents('xl/workbook.xml');

  # global regex match to gather pairs of names and ids
  my %sheets = ($contents =~ m[<sheet name="(.+?)" sheetId="(\d+)".*?>]g);

  return \%sheets;
}


#======================================================================
# METHODS
#======================================================================

sub values {
  my ($self, $sheet) = @_;
  my @data;
  my ($row, $col, $cell_type, $seen_node);

  my $contents = $self->_member_contents($self->sheet_member($sheet));

  while ($contents =~ m[<c\ r="([A-Z]+)(\d+)"      # initial cell tag; col and row
                            (?:[^>]*?t="(\w+)")?>  # maybe other attrs and cell type
                        (?:<v>([^<]+?)</v>         # cell value (if not inlineStr) ..
                        |                          # .. or
                           (.+?))                  # whole node content (if inlineStr)
                        </c>                       # closing cell tag
                        ]xg) {

    my ($col, $row, $cell_type, $val, $innerXML) = ($1, $2, $3, $4, $5);

    # convert column reference from A1 format to number format
    $col = $self->A1_to_num($col); 

    # handle cell value according to cell type
    $cell_type //= '';
    if ($cell_type eq 'inlineStr') {
      # this is an inline string; gather all <t> nodes within the cell node
      $val = join "", ($innerXML =~ m[<t>(.+?)</t>]g);
      $val =~ s/$regex_entities/$xml_entities{$1}/eg;
    }
    elsif ($cell_type eq 's') {
      # this is a string cell; $val is a pointer into the global array of shared strings
      $val = $self->strings->[$val];
    }
    elsif (! defined $val) {
      # try to find the <v> node within $innerXML -- maybe after a formula node
      ($val) = ($innerXML =~ m[<v>(.*?)</v>]);
      $val =~ s/$regex_entities/$xml_entities{$1}/eg if $val;
    }

    # insert this value into the global data array
    $data[$row-1][$col-1] = $val;
  }

  # insert arrayrefs for empty rows
  $_ //= [] foreach @data;

  return \@data;
}


1;

__END__

