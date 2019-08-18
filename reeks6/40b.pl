use bind_object;

@ARGV = ("ACS-Resource-Limits");
my $RootDSE = bind_object::bind_object("RootDSE");
my $LdapDisplayName = bind_object::bind_object("CN=" . $ARGV[0] . "," . $RootDSE->Get(schemaNamingContext) )->{ldapDisplayName};
@ARGV = ("$LdapDisplayName");
my $argument=$ARGV[0];

my $abstracteKlasse  = bind_object::bind_object( "schema/$argument" );
@attributen = qw(OID AuxDerivedFrom Abstract Auxiliary PossibleSuperiors MandatoryProperties
                 OptionalProperties Container Containment);

foreach my $prefix (@attributen){
    my $attribuut = $abstracteKlasse->{$prefix};
    printlijn( \$prefix, $_ ) foreach ref $attribuut eq "ARRAY" ? @{$attribuut} : $attribuut;
}

sub printlijn {
    my ( $refprefix, $suffix ) = @_;
    printf "%-35s%s\n", ${$refprefix}, $suffix;
    ${$refprefix} = "";
}