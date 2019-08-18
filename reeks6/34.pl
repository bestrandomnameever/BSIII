use bind_object;
use Data::Dumper;
use Win32::OLE::Variant;

my $RootDSE = bind_object::bind_object("RootDSE");
my $schemaNamingContext = $RootDSE->get(schemaNamingContext);

my $Schema = bind_object::bind_object($schemaNamingContext);

# Enkel klassen
$Schema->{Filter} = "classSchema";

push @classes, $_ foreach ( in $Schema);

# De naam van de properties heb ik opgezocht in de documentatie zoals er stond

# Alleen alle klasses die direct afgeleid zijn van top
@classes = grep {$_->{subClassOf} eq "top"} @classes;
print Dumper @classes;

# Enkel hulpklassen
# 3 = Auxiliary, dit kan je zien bij de documentatie
@classes = grep {$_->{objectClassCategory} eq 3} @classes;
print Dumper @classes;
 # De klasse kan attributen overnemen van één of meer hulpklassen
@classes = grep { defined $_->{auxiliaryClass} || defined $_->{systemAuxiliaryClass}} @classes;
print Dumper @classes;

# De klasse is Microsoft specifiek
@classes = grep { $_->{governsID} =~ /.*113556.*/} @classes;
print Dumper @classes;

# De klasse kan niet gewijzigd worden
@classes = grep { $_->{systemOnly} } @classes;
print Dumper @classes;

# De objecten van de klasse hebben een RDN die niet van de vorm CN=... is
@classes = grep { $_->{rDnAttId} ne "cn" } @classes;
print Dumper @classes;

foreach (@classes){
    print Dumper $_->{Name};
}