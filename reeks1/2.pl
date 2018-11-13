use strict;
# Waarom staat hier op het einde een array met in & with in?
# Omdat we 'submodules' willen laden van Win32::OLE. 
# in is nodig om een object te overlopen
# with is niet per se nodig zegt ze 
use Win32::OLE qw(in with);

while( (my $key, my $value) = each %INC ){
    print "\$INC{$key} -> $value\n";
}