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
use XML::LibXML::Reader qw(:types);


our @ISA     = qw(Excel::Reader::XLSX::Package::XMLreader);
our $VERSION = '0.00';

sub DEBUG {0};
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
    $self->{last_row}             = 0;

    bless $self, $class;

    return $self;
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
    $self->read_sheet_param() unless exists $self->{_init_worksh};
    
    return if $self->{empty};
    #for page_sheet_param 
    $self->{flag} = 1;
    # Read the next "row" element in the file.
    return if($self->{last_row} ==1);
    
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
    
    if ($row_number == $self->{p}{_range}[1][0]){
        $self->{last_row} = 1;
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

sub read_sheet_param{
    my $self = shift;
    
    $self->{flag} = 1 unless defined $self->{flag};
    
    while ( $self->{_reader}->read() && $self->{flag}) {
        $self->_parce_param( $self->{_reader} );
    }
   
}

sub _parce_param{
    my $self = shift;
    my $node = shift;

    # Only process the start elements.
    return unless $node->nodeType() == XML_READER_TYPE_ELEMENT;
    return if  $node->depth() != 1;
    if ($node->name eq 'dimension') {
        
        my ($r1,$r2)= split(/\:/,$self->{_reader}->getAttribute('ref'));
        my ($row,$col) = xl_cell_to_rowcol($r1);
        push @{$self->{p}{_range}},[$row,$col];
        $r2 = $r1 unless defined $r2;
        ($row,$col) = xl_cell_to_rowcol($r2);
        push @{$self->{p}{_range}},[$row,$col];
        $self->{_init_worksh} = 1;
        
    }elsif($node->name  eq 'cols'){
        while($self->{_reader}->read()){
            last unless ($self->{_reader}->name() eq "col");
            #for $worksheet->set_column( @{$self->{p}{_colAttr}});
            push @{$self->{p}{_colAttr}},$self->{_reader}->getAttribute('min');
            push @{$self->{p}{_colAttr}},$self->{_reader}->getAttribute('max');
            push @{$self->{p}{_colAttr}},$self->{_reader}->getAttribute('width'); 
       }
    }elsif($node->name  eq 'sheetData'){
        my $data = $node->isEmptyElement;
        if (!$node->isEmptyElement) {
            $self->{flag} =0;
        }else{
            $self->{empty} = 1;
        }  
        
    }elsif ( $node->name eq 'mergeCells' ) {
        $self->{_mergedcount} = $self->{_reader}->getAttribute( 'count' );
        while($self->{_mergedcount}){
            $self->{_reader}->nextElement();
            $self->{_mergedcount}--;
            push @{$self->{p}{_merged}} ,$self->{_reader}->getAttribute( 'ref' );
        }
        delete $self->{_mergedcount};
        
    }elsif($node->name eq 'pageMargins'){
        
        my $r = $node->getAttribute( 'right');
        my $l = $node->getAttribute('left');
        if ($r ==$l) {
            $self->{p}{_page}{margins_LR} = $r;
        }else{
            $self->{p}{_page}{margins_left} = $l;
            $self->{p}{_page}{margins_right} = $r;
        }
        $r = $node->getAttribute( 'top');
        $l = $node->getAttribute( 'bottom');
        if ($r ==$l) {
            $self->{p}{_page}{margins_TB} = $r;
        }else{
            $self->{p}{_page}{margins_top} = $r;
            $self->{p}{_page}{margins_bottom} = $l;
        }
        $self->{p}{_page}{headerfooter} = [$node->getAttribute( 'header'),$node->getAttribute( 'footer')];
        
    }elsif($node->name eq 'pageSetup'){
        #todo
        
    }elsif($node->name eq 'printOptions'){
        #$worksheet->center_horizontally();
        if ($node->moveToFirstAttribute){
            while (1) {
                $node->name() =~ /(.*)Cent.*/;
                $self->{p}{_print}{$1} =1;
                last unless $node->moveToNextAttribute;
            }
        }else{
            DEBUG && print "have printOptions but have no attribute in ",$self->{_name},"\n"  ;
        }

    }elsif($node->name eq 'headerFooter'){
        
        while (1) {
            $node->read();
            last if $node->name eq 'headerFooter';
            next unless $node->nodeType == XML_READER_TYPE_ELEMENT;
            $node->name =~ /odd(.*)/;
            $self->{p}{_page}{lc $1} = $node->readInnerXml;
            print $node->nodeType,"--",$node->name,"--",$node->value,"--",$node->depth,"\n";
        }
        
    }elsif($node->name eq 'rowBreaks'){
        # brk min and max ?
        $self->{_vcount} = $self->{_reader}->getAttribute( 'count' );
        while($self->{_vcount}){
            $self->{_reader}->nextElement();
            $self->{_vcount}--;
            push @{$self->{p}{_v_pagebreaks}} ,$self->{_reader}->getAttribute( 'id' );
        }
        delete $self->{_vcount};        
        
    }elsif($node->name eq 'colBreaks'){
        # brk min and max ?
        $self->{_hcount} = $self->{_reader}->getAttribute( 'count' );
        while($self->{_hcount}){
            $self->{_reader}->nextElement();
            $self->{_hcount}--;
            push @{$self->{p}{_h_pagebreaks}} ,$self->{_reader}->getAttribute( 'id' );
        }
        delete $self->{_hcount};
        
        
    }elsif($node->name eq 'sheetFormatPr' || $node->name eq 'sheetViews'){
        #i don't understand how use it
    }else{
        DEBUG && print "unknown sheet element: ",$node->name," in ",$self->{_name},"\n"  ;
        
    }
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

=head1 DESCRIPTION~

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