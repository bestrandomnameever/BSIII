use bind_object;
use Data::Dumper;
use Win32::OLE::Variant;

my $RootDSE = bind_object::bind_object("RootDSE");
$Schema = bind_object::bind_object($RootDSE->get(schemaNamingContext));


# Ik heb geen idee waar ik de ldapDisplayName van een bepaalde klasse zomaar kan vinden
# dus ik ga die van 37 eerst ophalen en daar dan de ldapDisplayName van gaan ophalen

# We hebben die ldapDisplayName nodig om te kunnen binden in het abstract schema
# De ldapdisplayname kan je bij het reele schema vinden in de distinguishedName

# VRAAG DIT VRIJDAG


# Dezelfde klasse als bij 37 gaan gebruiken
@ARGV = ("ACS-Resource-Limits");
my $LdapDisplayName = bind_object::bind_object("CN=" . $ARGV[0] . "," . $RootDSE->Get(schemaNamingContext) )->{ldapDisplayName};

$Object = bind_object::bind_object("schema/$LdapDisplayName");

@attributen = qw(OID AuxDerivedFrom Abstract Auxiliary PossibleSuperiors MandatoryProperties
                 OptionalProperties Container Containment);

foreach (@attributen){
    printf("%-20s: %-20s\n", $_, $Object->{$_} =~ /ARRAY.*/i ? join " ,", @{$Object->{$_}} : $Object->{$_});
}




