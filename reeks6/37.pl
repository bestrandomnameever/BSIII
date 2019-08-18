use bind_object;
use Data::Dumper;

# Een belangrijk probleem hierbij is dat de 
#distinguishedName van een schemaobject de RDN van het schemaobject bevat, 
#en niet de ldapDisplayName.

# Ben nog niet zo goed mee maar RDN = relative distinguished name ==> partial path to the
# entry relative to another entry in the tree

# Ik denk dat ze bedoelen dat er nog CN in de naam moet staan ofzo? 

# Foutmeldinghash maken
my %E_ADS = (
    BAD_PATHNAME            => Win32::OLE::HRESULT(0x80005000),
    UNKNOWN_OBJECT          => Win32::OLE::HRESULT(0x80005004),
    PROPERTY_NOT_SET        => Win32::OLE::HRESULT(0x80005005),
    PROPERTY_INVALID        => Win32::OLE::HRESULT(0x80005007),
    BAD_PARAMETER           => Win32::OLE::HRESULT(0x80005008),
    OBJECT_UNBOUND          => Win32::OLE::HRESULT(0x80005009),
    PROPERTY_MODIFIED       => Win32::OLE::HRESULT(0x8000500B),
    OBJECT_EXISTS           => Win32::OLE::HRESULT(0x8000500E),
    SCHEMA_VIOLATION        => Win32::OLE::HRESULT(0x8000500F),
    COLUMN_NOT_SET          => Win32::OLE::HRESULT(0x80005010),
    ERRORSOCCURRED          => Win32::OLE::HRESULT(0x00005011),
    NOMORE_ROWS             => Win32::OLE::HRESULT(0x00005012),
    NOMORE_COLUMNS          => Win32::OLE::HRESULT(0x00005013),
    INVALID_FILTER          => Win32::OLE::HRESULT(0x80005014),
    INVALID_DOMAIN_OBJECT   => Win32::OLE::HRESULT(0x80005001),
    INVALID_USER_OBJECT     => Win32::OLE::HRESULT(0x80005002),
    INVALID_COMPUTER_OBJECT => Win32::OLE::HRESULT(0x80005003),
    PROPERTY_NOT_SUPPORTED  => Win32::OLE::HRESULT(0x80005006),
    PROPERTY_NOT_MODIFIED   => Win32::OLE::HRESULT(0x8000500A),
    CANT_CONVERT_DATATYPE   => Win32::OLE::HRESULT(0x8000500C),
    PROPERTY_NOT_FOUND      => Win32::OLE::HRESULT(0x8000500D) );


my $RootDSE = bind_object::bind_object("RootDSE");
print Dumper $RootDSE;
#@ARGV = ("Account-Expires"); # Een attribuut
@ARGV = ("ACS-Resource-Limits");

# kommatje niet vergeten
my $Object = bind_object::bind_object("CN=" . $ARGV[0] . "," . $RootDSE->Get(schemaNamingContext) );

if ( $Object->{"Class"} eq "attributeSchema" ) {
    @attributen = qw (cn distinguishedName canonicalName ldapDisplayName
        attributeID attributeSyntax rangeLower rangeUpper
        isSingleValued isMemberOfPartialAttributeSet
        searchFlags  systemFlags);
    }

elsif ( $Object->{"Class"} eq "classSchema" ) {
    @attributen = qw(cn distinguishedName canonicalName ldapDisplayName
        governsID subClassOf systemAuxiliaryClass AuxiliaryClass
        objectClassCategory systemPossSuperiors possSuperiors
        systemMustContain mustContain systemMayContain mayContain);
    }
else { die("cn=$argument niet gevonden in  reeel schema\n"); }

# Attributen expliciet gaan opvragen
$Object->GetInfoEx([@attributen], 0);

foreach my $attribuut (@attributen) {
    my $prefix = $attribuut;

    # Tabel = referentie naar de array van values
    my $tabel  = $Object->GetEx($attribuut);

    # Getex smijt een error als het element niet gevonden is
    if(Win32::OLE->LastError() == $E_ADS{PROPERTY_NOT_FOUND}){
        printlijn( \$prefix, " < niet ingesteld > ");
    }
    else{
        printlijn( \$prefix, $_ ) foreach @{$tabel};
    }
}


sub printlijn {
    my ( $refprefix, $suffix ) = @_;
    printf "%-35s%s\n", ${$refprefix}, $suffix;
    ${$refprefix} = "";
}

