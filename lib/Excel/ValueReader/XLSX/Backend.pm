package Excel::ValueReader::XLSX::Backend;
use utf8;
use 5.10.1;
use Moose;
use Archive::Zip          qw(AZ_OK);

our $VERSION = '1.01';

#======================================================================
# ATTRIBUTES
#======================================================================
has 'frontend'      => (is => 'ro',   isa => 'Excel::ValueReader::XLSX', 
                        required => 1, weak_ref => 1,
                        handles => [qw/A1_to_num formatted_date/]);

has 'zip'           => (is => 'ro',   isa => 'Archive::Zip', init_arg => undef,
                        builder => '_zip', lazy => 1);

has 'date_styles'   => (is => 'ro',   isa => 'ArrayRef', init_arg => undef,
                        builder => '_date_styles', lazy => 1);

has 'strings'       => (is => 'ro',   isa => 'ArrayRef', init_arg => undef,
                        builder => '_strings',   lazy => 1);

has 'workbook_data' => (is => 'ro',   isa => 'HashRef', init_arg => undef,
                        builder => '_workbook_data',   lazy => 1);



#======================================================================
# ATTRIBUTE CONSTRUCTORS
#======================================================================



sub _zip {
  my $self = shift;

  my $xlsx_file = $self->frontend->xlsx;
  my $zip       = Archive::Zip->new;
  my $result    = $zip->read($xlsx_file);
  $result == AZ_OK  or die "cannot unzip $xlsx_file";

  return $zip;
}


#======================================================================
# METHODS
#======================================================================


sub base_year {
  my ($self) = @_;
  return $self->workbook_data->{base_year};
}

sub sheets {
  my ($self) = @_;
  return $self->workbook_data->{sheets};
}



sub Excel_builtin_date_formats {
  my @numFmt;

  # source : section 18.8.30 numFmt (Number Format) in ECMA-376-1:2016
  # Office Open XML File Formats — Fundamentals and Markup Language
  # Reference
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

  return @numFmt;
}


sub _zip_member_contents {
  my ($self, $member) = @_;

  my $contents = $self->zip->contents($member)
    or die "no contents for member $member";
  utf8::decode($contents);

  return $contents;
}



sub _zip_member_name_for_sheet {
  my ($self, $sheet) = @_;

  # check that sheet name was given
  $sheet or die "->values(): missing sheet name";

  # get sheet id
  my $id = $self->sheets->{$sheet};
  $id //= $sheet if $sheet =~ /^\d+$/;
  $id or die "no such sheet: $sheet";

  # construct member name for that sheet
  return "xl/worksheets/sheet$id.xml";
}






1;

__END__

=head1 NAME

Excel::ValueReader::XLSX::Backend -- TODO

=head1 DESCRIPTION


=head1 AUTHOR

Laurent Dami, E<lt>dami at cpan.orgE<gt>

=head1 COPYRIGHT AND LICENSE

Copyright 2021 by Laurent Dami.

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.