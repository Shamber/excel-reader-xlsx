package Excel::Reader::XLSX::Package::Styles;
use 5.008002;
use strict;
use warnings;
use Exporter;
use Carp;
use XML::LibXML::Reader qw(:types);
use Excel::Reader::XLSX::Package::XMLreader;

our @ISA     = qw(Excel::Reader::XLSX::Package::XMLreader);
our $VERSION = '0.00';

our $FULL_DEPTH  = 1;
our $RICH_STRING = 1;


###############################################################################
#
# new()
#
# Constructor.
#
sub new {

    my $class = shift;
    my $self  = Excel::Reader::XLSX::Package::XMLreader->new();

    $self->{_cnt}        = 0;

    bless $self, $class;

    return $self;
}


##############################################################################
#
# _read_all_nodes()
#
# Override callback function. TODO rename.
#
sub _read_all_nodes {

    my $self = shift;
    my $reader = $self->{_reader};
    
    while ($reader->read()) {
        $self->_read_node($reader);
    }
}

sub _read_node {

    my $self = shift;
    my $node = shift;

    # Only process the start elements.
    return unless $node->nodeType() == XML_READER_TYPE_ELEMENT;
    
    $self->{_cnt} = $self->{_reader}->getAttribute('count');
    
    if ( $node->name eq 'numFmts' ) {
        return if $self->{_cnt} == 1;
        while($self->_next_Element()){
            my $f = $self->{_reader}->getAttribute('numFmtId');
            $self->{_numFmt}{$f} = $self->{_reader}->getAttribute('formatCode');
        }
    }elsif($node->name eq 'fonts'){
        $self->_parce_param('_font','val');
    }elsif($node->name eq 'fills'){
        $self->_parce_param('_fill','patternType');
    }elsif($node->name eq 'borders'){
        $self->_parce_param('_border','style');
    }elsif($node->name eq 'cellXfs'){
        #
  
    }elsif($node->name eq 'extLst' or $node->name eq 'cellStyles' or $node->name eq 'cellStyleXfs' ){
        
    }elsif($node->name eq 'colors' or $node->name eq 'tableStyles' or $node->name eq 'dxfs'){
        
    }
}


sub _parce_param{
    my $self =shift;
    my $param = shift;
    my $attrname = shift;
    
    return if $self->{_cnt} == 1;
    while($self->_next_Element()){
        my $ret;
        my $depth = $self->{_reader}->depth;
        while ($self->{_reader}->read()) {
            last if ($self->{_reader}->depth == $depth);
            my $n = $self->{_reader}->name;
            my $val = $self->{_reader}->getAttribute($attrname);
            $val = $val || " ";
            #check if this first font
            unless (exists $self->{$param}) {
            #save default parametr
                $ret->{$n} = $val;
            }else{
                if (exists $self->{$param}->[0]->{$n}){
                    #save only changed parametr
                    $ret->{$n} = $val unless $self->{$param}[0]->{$n} eq $val;
                }else{
                    $ret->{$n} = $val;
                }
            }
        }
        $ret = $ret || 'none';
        push @{$self->{$param}},$ret;
    }
    
}

sub _next_Element{
    my $self =shift;
    
    if ($self->{_cnt}){
        $self->{_reader}->nextElement();
        $self->{_cnt}--;
        return 1;
    }else{
        delete $self->{_cnt};
        return 0;
    }
    
}

1;


__END__