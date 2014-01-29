package Excel::Reader::XLSX::Worksheet;

###############################################################################
#
# Worksheet - A class for reading the Excel XLSX sheet.xml file.
#
# Used in conjunction with Excel::Reader::XLSX
#
# Copyright 2012, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

# perltidy with the following options: -mbl=2 -pt=0 -nola

use 5.008002;
use strict;
use warnings;
use Carp;
use Excel::Reader::XLSX::Package::XMLreader;
use Excel::Reader::XLSX::Row;


our @ISA     = qw(Excel::Reader::XLSX::Package::XMLreader);
our $VERSION = '0.00';

###############################################################################
#
# Public and private API methods.
#
###############################################################################


###############################################################################
#
# new()
#
# Constructor.
#
sub new {

    my $class = shift;
    my $self  = Excel::Reader::XLSX::Package::XMLreader->new();

    $self->{_shared_strings}      = shift;
    $self->{_name}                = shift;
    $self->{_index}               = shift;
    $self->{_previous_row_number} = -1;

    bless $self, $class;

    return $self;
}

###############################################################################
#
# _init_worksheet()
# set $self->{_range} cell range
# set $self->{width} for col
# set $self->{sheetview} for col
# TODO.
#

sub _init_worksheet{
    my $self = shift;
    if ($self->{_reader}->nextElement( 'dimension' )){
        my ($r1,$r2)= split(/\:/,$self->{_reader}->getAttribute('ref'));
        my ($row,$col) = xl_cell_to_rowcol($r1);
        push @{$self->{_range}},[$row,$col];
        ($row,$col) = xl_cell_to_rowcol($r2);
        push @{$self->{_range}},[$row,$col];
    }
    
    if ($self->{_reader}->nextElement('cols')) {
        while($self->{_reader}->read()){
            last unless ($self->{_reader}->name() eq "col");
            #for $worksheet->set_column( @{$self->{_colAttr}});
            push @{$self->{_colAttr}},$self->{_reader}->getAttribute('min');
            push @{$self->{_colAttr}},$self->{_reader}->getAttribute('max');
            push @{$self->{_colAttr}},$self->{_reader}->getAttribute('width'); 
       }
    }
    # send to first row
    $self->DESTROY() unless $self->{_reader}->nextElement('row');
    $self->{_init_worksh} = 1;
    
}
###############################################################################
#
# _init_row()
#
# TODO.
#
sub _init_row {

    my $self = shift;

    # Store reusable Cell object to avoid repeated calls to Cell::new().
    $self->{_cell} = Excel::Reader::XLSX::Cell->new( $self->{_shared_strings} );

    # Store reusable Row object to avoid repeated calls to Row::new().
    $self->{_row}       = Excel::Reader::XLSX::Row->new(
        $self->{_reader},
        $self->{_shared_strings},
        $self->{_cell},
    );
    $self->{_row_initialised} = 1;
}


###############################################################################
#
# next_row()
#
# Read the next available row in the worksheet.
#
sub next_row {

    my $self = shift;
    my $row  = undef;

    #Read  dimension and col width
    $self->_init_worksheet() unless exists $self->{_init_worksh};
    
    # Read the next "row" element in the file.   
    return unless $self->{_reader}->name() eq "row";
    # Read the row attributes.
    my $row_reader = $self->{_reader};
    my $row_number = $row_reader->getAttribute( 'r' );

    # Zero index the row number.
    if ( defined $row_number ) {
        $row_number--;
    }
    else {
        # If no 'r' attribute assume it is one more than the previous.
        $row_number = $self->{_previous_row_number} + 1;
    }


    if ( !$self->{_row_initialised} ) {
        $self->_init_row();
    }

    $row = $self->{_row};
    $row->_init( $row_number, $self->{_previous_row_number}, );

    $self->{_previous_row_number} = $row_number;
    
    if ($row_number == $self->{_range}[1][0]){
        $self->{_reader}->nextElement() 
    }else{
        $self->{_reader}->nextElement('row');
    }
    return $row;
}
###############################################################################
#
# name()
#
# Return the worksheet name.
#
sub name {
    my $self = shift;
    return $self->{_name};
}

###############################################################################
#
# index()
#
# Return the worksheet index.
#
sub index {
    my $self = shift;
    return $self->{_index};
}

sub xl_cell_to_rowcol {
    my $cell = shift;
    return ( 0, 0, 0, 0 ) unless $cell;
    $cell =~ /(\$?)([A-Z]{1,3})(\$?)(\d+)/;
    my $col_abs = $1 eq "" ? 0 : 1;
    my $col     = $2;
    my $row_abs = $3 eq "" ? 0 : 1;
    my $row     = $4;

    # Convert base26 column string to number
    # All your Base are belong to us.
    my @chars = split //, $col;
    my $expn = 0;
    $col = 0;

    while ( @chars ) {
        my $char = pop( @chars );    # LS char first
        $col += ( ord( $char ) - ord( 'A' ) + 1 ) * ( 26**$expn );
        $expn++;
    }

    # Convert 1-index to zero-index
    $row--;
    $col--;

    return $row, $col;
}



sub merged{
    my $self = shift;
    return unless $self->{_reader}->nextElement( 'mergeCells' );
    my $merged_row = $self->{_reader};
    $self->{_mergedcount} = $merged_row->getAttribute( 'count' );
    while( $merged_row->nextElement){
        push @{$self->{_merged}} ,$merged_row->getAttribute( 'ref' );
    }
   return 1;
}

###############################################################################
#
# Internal methods.
#
###############################################################################


1;


__END__

=pod

=head1 NAME

Worksheet - A class for reading the Excel XLSX sheet.xml file.

=head1 SYNOPSIS

See the documentation for L<Excel::Reader::XLSX>.

=head1 DESCRIPTION

This module is used in conjunction with L<Excel::Reader::XLSX>.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

Copyright MMXII, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

=head1 LICENSE

Either the Perl Artistic Licence L<http://dev.perl.org/licenses/artistic.html> or the GPL L<http://www.opensource.org/licenses/gpl-license.php>.

=head1 DISCLAIMER OF WARRANTY

See the documentation for L<Excel::Reader::XLSX>.

=cut