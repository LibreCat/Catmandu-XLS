package Catmandu::Importer::XLSX;

use namespace::clean;
use Catmandu::Sane;
use Encode qw(decode);
use Spreadsheet::XLSX;
use Spreadsheet::ParseExcel::Utility qw(int2col);
use Moo;

with 'Catmandu::Importer';

has xlsx => (is => 'ro', builder => '_build_xlsx');
has header => (is => 'ro', default => sub { 1 });
has columns => (is => 'ro' , default => sub { 0 });
has fields => (
    is     => 'rw',
    coerce => sub {
        my $fields = $_[0];
        if (ref $fields eq 'ARRAY') { return $fields }
        if (ref $fields eq 'HASH')  { return [sort keys %$fields] }
        return [split ',', $fields];
    },
);
has _n => (is => 'rw', default => sub { 0 });
has _row_min => (is => 'rw');
has _row_max => (is => 'rw');
has _col_min => (is => 'rw');
has _col_max => (is => 'rw');

sub BUILD {
    my $self = shift;

    if ( $self->header ) {
        if ( $self->fields ) {
            $self->{_n}++;
        }
        elsif ( $self->columns ) {
            $self->fields([$self->_get_cols]);
            $self->{_n}++;
        }
        else {
            $self->fields([$self->_get_row]);
            $self->{_n}++;
        }
    }
    else {
        if ( !$self->fields || $self->columns ) {
            $self->fields([$self->_get_cols]);
        }
    }
}

sub _build_xlsx {
    my ($self) = @_;
    my $xlsx = Spreadsheet::XLSX->new( $self->file ) or die "Could not open file " . $self->file;

    # process only first worksheet
    $xlsx = $xlsx->{Worksheet}->[0];
    $self->{_col_min} = $xlsx->{MinCol};
    $self->{_col_max} = $xlsx->{MaxCol};
    $self->{_row_min} = $xlsx->{MinRow};
    $self->{_row_max} = $xlsx->{MaxRow};
    return $xlsx;
}

sub generator {
    my ($self) = @_;
    sub {
        while ($self->_n <= $self->_row_max) {
            my @data = $self->_get_row();
            $self->{_n}++;
            my @fields = @{$self->fields()};
            my %hash = map { 
                my $key = shift @fields;
                defined $_  ? ($key => $_) : ()
                } @data;
            return \%hash;
        }
        return;
    }
}

sub _get_row {
    my ($self) = @_;
    my @row;
    for my $col ( $self->_col_min .. $self->_col_max ) {
        my $cell = $self->xlsx->{Cells}[$self->_n][$col];
        if ($cell) {
            push(@row, decode('UTF-8',$cell->{Val}));
        }
        else{
            push(@row, undef);            
        }
    }
    return @row;
}

sub _get_cols {
    my ($self) = @_;
    my @row;
    for my $col ( $self->_col_min .. $self->_col_max ) {

        if (!$self->header || $self->columns) {
            push(@row,int2col($col));
        }
        else {
            my $cell = $self->xlsx->{Cells}[$self->_n][$col];
            if ($cell) {
                push(@row, decode('UTF-8',$cell->{Val}));
            }
            else{
                push(@row, undef);            
            }
        }
    }
    return @row;
}


=head1 NAME

Catmandu::Importer::XLSX - Package that imports XLSX files

=head1 SYNOPSIS

    # On the command line
    $ catmandu convert XLSX < ./t/test.xlsx
    $ catmandu convert XLSX --header 0 < ./t/test.xlsx
    $ catmandu convert XLSX --fields 1,2,3 < ./t/test.xlsx
    $ catmandu convert XLSX --columns 1 < ./t/test.xlsx

    # Or in Perl
    use Catmandu::Importer::XLSX;

    my $importer = Catmandu::Importer::XLSX->new(file => "./t/test.xls");

    my $n = $importer->each(sub {
        my $hashref = $_[0];
        # ...
    });

=head1 DESCRIPTION

L<Catmandu> importer for XLSX files.

Only the first worksheet from the Excel workbook is imported.

=head1 METHODS
 
This module inherits all methods of L<Catmandu::Importer> and by this
L<Catmandu::Iterable>.
 
=head1 CONFIGURATION
 
In addition to the configuration provided by L<Catmandu::Importer> (C<file>,
C<fh>, etc.) the importer can be configured with the following parameters:
 
=over
 
=item header

By default object fields are read from the XLS header line. If no header 
line is avaiable object fields are named as column coordinates (A,B,C,...). Default: 1.

=item fields

Provide custom object field names as array, hash reference or comma-
separated list.

=item columns

When the 'columns' option is provided, then the object fields are named as 
column coordinates (A,B,C,...). Default: 0.
 
=back

=head1 SEE ALSO

L<Catmandu::Importer>, L<Catmandu::Iterable>, L<Catmandu::Importer::CSV>, L<Catmandu::Importer::XLS>.

=cut

1;