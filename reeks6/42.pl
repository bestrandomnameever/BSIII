use bind_object;
use Data::Dumper;
use Win32::OLE::Variant;

my $RootDSE = bind_object::bind_object("RootDSE");
# Filter direct zodat je enkel attributen krijgt.
my $Schema = bind_object::bind_object($RootDSE->Get("schemaNamingContext"));
$Schema->{Filter} = "attributeSchema";

foreach(in $Schema){
    # abstract attribuut gaan ophalen zodat je de syntax kan krijgen
    # Je moet het gaan ophalen met ldapDisplayName
    my $abstractAttribute = bind_object::bind_object("schema/" . $_->{ldapDisplayName});
    
    # Er zijn blijkbaar 2 attributen die de syntax beschrijven van reeële attributen
    # attributeSyntax en omSyntax

    # hash vullen
    $map->{$abstractAttribute->{Syntax}}{attributeSyntax} = $_->{attributeSyntax};
    $map->{$abstractAttribute->{Syntax}}{omSyntax} = $_->{omSyntax};

}

print Dumper $map;
