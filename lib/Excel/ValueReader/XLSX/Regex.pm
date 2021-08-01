package Excel::ValueReader::XLSX::Regex;
use utf8;
use 5.10.1;
use Moose;
use Date::Calc qw/Add_Delta_Days/;

#======================================================================
# GLOBAL VARIABLES
#======================================================================

our $VERSION = '1.01';

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
                    handles => [qw/sheet_member _member_contents strings A1_to_num
                                   base_year date_formatter/]);

has 'date_styles' => (is => 'ro',   isa => 'ArrayRef', init_arg => undef,
                      builder => '_date_styles', lazy => 1);

#======================================================================
# LAZY ATTRIBUTE CONSTRUCTORS
#======================================================================

sub _strings {
  my $self = shift;
  my @strings;

  # read from the sharedStrings zip member
  my $contents = $self->_member_contents('xl/sharedStrings.xml');

  # iterate on <si> nodes
  while ($contents =~ m[<si>(.*?)</si>]sg) {
    my $innerXML = $1;

    # concatenate contents from all <t> nodes (usually there is only 1)
    my $string   = join "", ($innerXML =~ m[<t[^>]*>(.+?)</t>]sg);

    # decode entities
    $string =~ s/$regex_entities/$xml_entities{$1}/eg;

    push @strings, $string;
  }

  return \@strings;
}


sub _workbook_data {
  my $self = shift;

  # read from the workbook.xml zip member
  my $workbook = $self->_member_contents('xl/workbook.xml');

  # extract sheet names
  my @sheet_names = ($workbook =~ m[<sheet name="(.+?)"]g);
  my %sheets      = map {$sheet_names[$_] => $_+1} 0 .. $#sheet_names;

  # does this workbook use the 1904 calendar ?
  my ($date1904) = $workbook =~ m[date1904="(.+?)"];
  my $base_year  = $date1904 ? 1904 : 1900;

  return {sheets => \%sheets, base_year => $base_year};
}



sub _date_styles {
  my $self = shift;

  state $date_style_regex = qr{[dy]};

  # read from the styles.xml zip member
  my $styles = $self->_member_contents('xl/styles.xml');

  # start with Excel builtin number formats for dates and times
  my @numFmt;
  $numFmt[14] = 'mm-dd-yy';
  $numFmt[15] = 'd-mmm-yy';
  $numFmt[16] = 'd-mmm';
  $numFmt[17] = 'mmm-yy';
  # $numFmt[18] = 'h:mm AM/PM';
  # $numFmt[19] = 'h:mm:ss AM/PM';
  # $numFmt[20] = 'h:mm';
  # $numFmt[21] = 'h:mm:ss';
  $numFmt[22] = 'm/d/yy h:mm';
  # $numFmt[45] = 'mm:ss';
  # $numFmt[46] = '[h]:mm:ss';
  # $numFmt[47] = 'mmss.0';

  # other specific date formats specified in this workbook
  while ($styles =~ m[<numFmt numFmtId="(\d+)" formatCode="([^"]+)"/>]g) {
    my ($id, $code) = ($1, $2);
    $numFmt[$id] = $code if $code =~ $date_style_regex;
  }

  # read all cell formats, just rembember those that involve a date number format
  my ($cellXfs)    = ($styles =~ m[<cellXfs count="\d+">(.+?)</cellXfs>]);
  my @cell_formats = $self->_extract_xf($cellXfs);
  my @date_styles  = map {$numFmt[$_->{numFmtId}]} @cell_formats;

  return \@date_styles; # array of shape (xf_index => numFmt_code)
}

sub _extract_xf {
  my ($self, $xml) = @_;

  state $xf_node_regex = qr{
   <xf                  # initial format tag
     \s
     ([^>/]*+)          # attributes (captured in $1)
     (?:                # non-capturing group for an alternation :
        />              # .. either an xml closing without content
      |                 # or
        >               # .. closing for the xf tag
        .*?             # .. then some formatting content
       </xf>            # .. then the ending tag for the xf node
     )
    }x;

  my @xf_nodes;
  while ($xml =~ /$xf_node_regex/g) {
    my $all_attrs = $1;
    my %attr;
    while ($all_attrs =~ m[(\w+)="(.+?)"]g) {
      $attr{$1} = $2;
    }
    push @xf_nodes, \%attr;
  }
  return @xf_nodes;
}


#======================================================================
# METHODS
#======================================================================

sub values {
  my ($self, $sheet) = @_;
  my @data;
  my ($row, $col, $cell_type, $seen_node);

  state $cell_regex = qr(
     <c\                     # initial cell tag
      r="([A-Z]+)(\d+)"      # capture col and row ($1 and $2)
      [^>/]*?                # unused attrs
      (?:s="(\d+)"\s*)?      # style attribute ($3)
      (?:t="(\w+)"\s*)?      # type attribute ($4)
     (?:                     # non-capturing group for an alternation :
        />                   # .. either an xml closing without content
      |                      # or
        >                    # .. closing xml tag, followed by
      (?:

         <v>(.+?)</v>        #    .. a value ($5)
        |                    #    or 
          (.+?)              #    .. some node content ($6)
       )
       </c>                  #    followed by a closing cell tag
      )
    )x;
  # NOTE : this regex uses capturing groups; I tried with named captures
  # but this doubled the execution time on big Excel files

  # parse worksheet XML, gathering all cells
  my $contents = $self->_member_contents($self->sheet_member($sheet));
  while ($contents =~ /$cell_regex/g) {
    my ($col, $row, $style, $cell_type, $val, $inner) = ($self->A1_to_num($1), $2, $3, $4, $5, $6);

    # handle cell value according to cell type
    $cell_type //= '';
    if ($cell_type eq 'inlineStr') {
      # this is an inline string; gather all <t> nodes within the cell node
      $val = join "", ($inner =~ m[<t>(.+?)</t>]g);
      $val =~ s/$regex_entities/$xml_entities{$1}/eg if $val;
    }
    elsif ($cell_type eq 's') {
      # this is a string cell; $val is a pointer into the global array of shared strings
      $val = $self->strings->[$val];
    }
    else {
      ($val) = ($inner =~ m[<v>(.*?)</v>])           if !defined $val && $inner;
      $val =~ s/$regex_entities/$xml_entities{$1}/eg if $val && $cell_type eq 'str';

      if ($style && defined $val && $val >= 0) {
        my $date_style = $self->date_styles->[$style];
        $val = $self->_formatted_date($val, $date_style)    if $date_style;
      }
    }

    # insert this value into the global data array
    $data[$row-1][$col-1] = $val;
  }

  # insert arrayrefs for empty rows
  $_ //= [] foreach @data;

  return \@data;
}



sub _formatted_date {
  my ($self, $val, $date_style) = @_;

  state $millisecond = 1 / (24*60*60*1000);

  my $n_days     = int($val);
  my $fractional = $val - $n_days;

  my $base_year  = $self->base_year;

  $n_days -= 1;                                       # because we need a 0-based value
  $n_days -=1 if $base_year == 1900 && $n_days >= 60; # Excel believes 1900 is a leap year

  my @d = Add_Delta_Days($base_year, 1, 1, $n_days);

  foreach my $subdivision (24, 60, 60, 1000) {
    last if abs($fractional) < $millisecond;
    $fractional *= $subdivision;
    my $unit = int($fractional);
    $fractional -= $unit;
    push @d, $unit;
  }
  
  # TEMP
  # return "$val($date_style)";

  return $self->date_formatter->(@d);

}


1;

__END__

=head1 NAME

Excel::ValueReader::XLSX::Regex - using regexes for extracting values from Excel workbooks

=head1 DESCRIPTION

This is one of two backend modules for L<Excel::ValueReader::XLSX>; the other
possible backend is L<Excel::ValueReader::XLSX::LibXML>.

This backend parses OOXML structures using regular expressions.

=head1 AUTHOR

Laurent Dami, E<lt>dami at cpan.orgE<gt>

=head1 COPYRIGHT AND LICENSE

Copyright 2020 by Laurent Dami.

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.
