package Excel::ValueReader::XLSX;
use utf8;
use Moose;
use Archive::Zip          qw(AZ_OK);
use Encode                qw(decode_utf8);
use Module::Load          qw/load/;
use feature 'state';

our $VERSION = '1.0';

#======================================================================
# ATTRIBUTES
#======================================================================

# public attributes
has 'xlsx'      => (is => 'ro', isa => 'Str', required => 1);
has 'using'     => (is => 'ro',   isa => 'Str', default => 'Regex');

# attributes used internally, not documented
has 'zip'       => (is => 'ro',   isa => 'Archive::Zip', init_arg => undef,
                    builder => '_zip',   lazy => 1);
has 'sheets'    => (is => 'ro',   isa => 'HashRef', init_arg => undef,
                    builder => '_sheets',   lazy => 1);
has 'strings'   => (is => 'ro',   isa => 'ArrayRef', init_arg => undef,
                    builder => '_strings',   lazy => 1);
has 'backend'   => (is => 'ro',   isa => 'Ref', 
                    builder => '_backend',   lazy => 1,
                    handles => [qw/_strings _sheets values/]);

#======================================================================
# BUILDING
#======================================================================

# syntactic sugar for supporting ->new($path) instead of ->new(docx => $path)
around BUILDARGS => sub {
  my $orig  = shift;
  my $class = shift;

  if ( @_ == 1 && !ref $_[0] ) {
    return $class->$orig(xlsx => $_[0]);
  }
  else {
    return $class->$orig(@_);
  }
};



#======================================================================
# LAZY ATTRIBUTE CONSTRUCTORS
#======================================================================

sub _zip {
  my $self = shift;

  my $zip = Archive::Zip->new;
  $zip->read($self->{xlsx}) == AZ_OK
      or die "cannot unzip $self->{docx}";

  return $zip;
}


sub _backend {
  my $self = shift;

  my $backend_class = ref($self) . '::' . $self->using;
  load $backend_class;

  return $backend_class->new(frontend => $self);
}



#======================================================================
# METHODS
#======================================================================

sub sheet_names {
  my ($self) = @_;

  my $sheets = $self->sheets;

  return sort {$sheets->{$a} <=> $sheets->{$b}} keys %$sheets;
}



sub _member_contents {
  my ($self, $member) = @_;

  my $bytes    = $self->zip->contents($member)
    or die "no contents for member $member";

  my $contents = decode_utf8($bytes);

  return $contents;
}



sub sheet_member {
  my ($self, $sheet) = @_;

  # check that sheet name was given
  $sheet or die "->values(): missing sheet name";

  # get sheet id
  my $id = $self->sheets->{$sheet};
  $id //= $sheet if $sheet =~ /^\d+$/;
  $id or die "no such sheet: $sheet";


  return "xl/worksheets/sheet$id.xml";
}


sub A1_to_num { # convert Excel A1 reference format to a number
  my ($self, $string) = @_;;

  state $base = ord('A') - 1;

  my $num = 0;
  foreach my $digit (map {ord($_) - $base} split //, $string) {
    $num = $num*26 + $digit;
  }

  return $num;
}





1;


=cut









1;

__END__

=head1 NAME

Excel::ValueReader::XLSX -- extracting values from Excel workbooks in XLSX format, fast

=head1 SYNOPSIS

  my $extractor = Excel::ValueReader::XLSX->new(xlsx => $filename);

  foreach my $sheet_name ($extractor->sheet_names) {
     my $grid = $extractor->values($sheet_name);
     my $n_rows = @$grid;
     print "sheet $sheet_name has $n_rows rows; ",
           "first cell contains : ", $grid->[0][0];
  }

=head1 DESCRIPTION

This module reads the contents of an Excel file in XLSX format,
and returns a bidimensional array of values for each worksheet.

Unlike L<Spreadsheet::ParseXLSX> or L<Spreadsheet::XLSX>, there is no
API to read formulas, formats or other Excel information; all you get
are plain values -- but you get them much faster than with these other modules !



=head1 METHODS

=head1 SEE ALSO

The official reference for OOXML-XLSX format is in
L<https://www.ecma-international.org/publications/standards/Ecma-376.htm>.

Introductory material can be found at
L<http://officeopenxml.com/anatomyofOOXML-xlsx.php>.

Another unpublished but working module for parsing Excel files in perl
can be found at L<https://github.com/jmcnamara/excel-reader-xlsx>.

