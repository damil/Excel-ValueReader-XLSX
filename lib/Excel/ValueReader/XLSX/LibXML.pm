package Excel::ValueReader::XLSX::LibXML;
use utf8;
use 5.10.1;
use Moose;
use XML::LibXML::Reader;
use Date::Calc qw/Add_Delta_Days/;

our $VERSION = '1.01';

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

  my $reader = $self->_reader_for_member('xl/sharedStrings.xml');

  my @strings;
  my $last_string = '';
  while ($reader->read) {
    my $node_name = $reader->name;

    if ($node_name eq 'si') {
      push @strings, $last_string if $last_string;
      $last_string = '';
    }
    elsif ($node_name eq '#text') {
      $last_string .= $reader->value;
    }
  }

  push @strings, $last_string if $last_string;

  return \@strings;
}


sub _workbook_data {
  my $self = shift;

  my %sheets;
  my $sheet_id  = 1;
  my $base_year = 1900;


  my $reader = $self->_reader_for_member('xl/workbook.xml');

  while ($reader->read) {
    if ($reader->name eq 'sheet' && $reader->nodeType == XML_READER_TYPE_ELEMENT) {
      my $name = $reader->getAttribute('name')
        or die "sheet node without name";
      $sheets{$name} = $sheet_id++;
    }
    elsif ($reader->name eq 'workbookPr' && $reader->getAttribute('date1904')) {
      $base_year = 1904; # this workbook uses the 1904 calendar
    }
  }

  return {sheets => \%sheets, base_year => $base_year};
}

sub _date_styles {
  my $self = shift;

  state $date_style_regex = qr{[dy]|\bmm\b};
  my @date_styles;

  # read from the styles.xml zip member
  my $reader = $self->_reader_for_member('xl/styles.xml');

  # start with Excel builtin number formats for dates and times
  my @numFmt;
  $numFmt[14] = 'mm-dd-yy';
  $numFmt[15] = 'd-mmm-yy';
  $numFmt[16] = 'd-mmm';
  $numFmt[17] = 'mmm-yy';
  $numFmt[18] = 'h:mm AM/PM';
  $numFmt[19] = 'h:mm:ss AM/PM';
  $numFmt[20] = 'h:mm';
  $numFmt[21] = 'h:mm:ss';
  $numFmt[22] = 'm/d/yy h:mm';
  $numFmt[45] = 'mm:ss';
  $numFmt[46] = '[h]:mm:ss';
  $numFmt[47] = 'mmss.0';

  my $expected_subnode = undef;

  # add other date formats explicitly specified in this workbook
 NODE:
  while ($reader->read) {

    $reader->nodeType == XML_READER_TYPE_ELEMENT or next NODE;


    # special treatment for some specific subtrees
    if ($expected_subnode) {
      my ($name, $depth, $handler) = @$expected_subnode;
      if ($reader->name eq $name && $reader->depth == $depth) {
        # process that subnode and go to the next node
        $handler->();
        next NODE;
      }
      elsif ($reader->depth < $depth) {
        # finished handling subnodes; back to regular node treatment
        $expected_subnode = undef;
      }
    }

    # regular node treatement
    if ($reader->name eq 'numFmts') {
      $expected_subnode = [numFmt => $reader->depth+1 => sub {
                             my $id   = $reader->getAttribute('numFmtId');
                             my $code = $reader->getAttribute('formatCode');
                             $numFmt[$id] = $code if $id && $code && $code =~ $date_style_regex;
                           }];
    }

    elsif ($reader->name eq 'cellXfs') {
      $expected_subnode = [xf => $reader->depth+1 => sub {
                             state $xf_count = 0;
                             my $numFmtId    = $reader->getAttribute('numFmtId');
                             my $code        = $numFmt[$numFmtId];
                             $date_styles[$xf_count++] = $code; # may be undef
                           }];
    }
  }

  return \@date_styles;
}



#======================================================================
# METHODS
#======================================================================

sub _reader_for_member {
  my ($self, $member) = @_;

  my $reader = XML::LibXML::Reader->new(string     => $self->_member_contents($member),
                                        no_blanks  => 1,
                                        no_network => 1,
                                        huge       => 1);
  return $reader;
}

sub values {
  my ($self, $sheet) = @_;

  # prepare for traversing the XML structure
  my $reader = $self->_reader_for_member($self->sheet_member($sheet));
  my @data;
  my ($row, $col, $cell_type, $cell_style, $seen_node);

  # iterate through XML nodes
  while ($reader->read) {
    my $node_name = $reader->name;

    if ($node_name eq 'c') {
      # new cell node : store its col/row reference and its type
      my $A1_cell_ref = $reader->getAttribute('r');
      ($col, $row)    = ($A1_cell_ref =~ /^([A-Z]+)(\d+)$/);
      $col            = $self->A1_to_num($col);
      $cell_type      = $reader->getAttribute('t');
      $cell_style     = $reader->getAttribute('s');
      $seen_node      = '';
    }

    elsif ($node_name =~ /^[vtf]$/) {
      # remember we have seen a 'value' or 'text' or 'formula' node
      $seen_node = $node_name;
    }

    elsif ($node_name eq '#text') {
      #start processing cell content

      my $val = $reader->value;
      $cell_type //= '';

      if ($seen_node eq 'v')  {
        if ($cell_type eq 's') {
          $val = $self->strings->[$val]; # string -- pointer into the global
                                         # array of shared strings
        }
        elsif ($cell_type eq 'e') {
          $val = undef; # error -- silently replace by undef
        }
        elsif ($cell_type =~ /^(n|d|b|str|)$/) {
          # number, date, boolean, formula string or no type : content is already in $val

          # if this is a date, replace the numeric value by the formatted date
          if ($cell_style && defined $val && $val >= 0) {
            my $date_style = $self->date_styles->[$cell_style];
            $val = $self->_formatted_date($val, $date_style)    if $date_style;
          }
        }
        else {
          # handle unexpected cases
          warn "unsupported type '$cell_type' in cell L${row}C${col}\n";
          $val = undef;
        }

        # insert this value into the global data array
        $data[$row-1][$col-1] = $val;
      }

      elsif ($seen_node eq 't' && $cell_type eq 'inlineStr')  {
        # inline string -- accumulate all #text nodes until next cell
        no warnings 'uninitialized';
        $data[$row-1][$col-1] .= $val;
      }

      elsif ($seen_node eq 'f')  {
        # formula -- just ignore it
      }

      else {
        # handle unexpected cases
        warn "unexpected text node in cell L${row}C${col}: $val\n";
      }
    }
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

  return $self->date_formatter->(@d);
}


1;


__END__


=head1 NAME

Excel::ValueReader::XLSX::LibXML - using LibXML for extracting values from Excel workbooks

=head1 DESCRIPTION

This is one of two backend modules for L<Excel::ValueReader::XLSX>; the other
possible backend is L<Excel::ValueReader::XLSX::Regex>.

This backend parses OOXML structures using L<XML::LibXML::Reader>.

=head1 AUTHOR

Laurent Dami, E<lt>dami at cpan.orgE<gt>

=head1 COPYRIGHT AND LICENSE

Copyright 2020 by Laurent Dami.

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.
