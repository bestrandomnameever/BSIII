use bind_object;
use Data::Dumper;
use Win32::OLE::Variant;

my $RootDSE = bind_object::bind_object("RootDSE");
my $schemaNamingContext = $RootDSE->get(schemaNamingContext);

my $Schema = bind_object::bind_object($schemaNamingContext);

# Enkel attributen
$Schema->{Filter} = "attributeSchema";

push @attributes, $_ foreach (in $Schema);

# is multivalued
@attributes = grep {!($_->{isSingleValued})} @attributes;
print Dumper @attributes;

# is opgenomen in de Global Catalog
@attributes = grep {$_->{isMemberOfPartialAttributeSet}} @attributes;

# het attribuut laat indexering toe

# An attributes is indexed when the searchFlags attribute in the attribute's schema definition 
# has the least significant bit set to 1. Setting the least significant bit of the searchFlags 
# attribute schema definition to 1 will dynamically build an index. Setting the least significant 
# bit of the searchFlags attribute schema definition to 0 will cause the index for the attribute to be removed. 
# The index will be built automatically by a background thread on the domain controller.

@attributes = grep {$_->{searchFlags} & 1} @attributes;

# Het attribuut wordt niet gerepliceerd tussen domeincontrollers
# Kijk in de systemflags
@attributes = grep {$_->{systemflags} & 1} @attributes;

# het LDAP-attribuut is geconstrueerd op basis van andere attributen
@attributes = grep {$_->{systemflags} & 4} @attributes;

# het LDAP-attribuut heeft een ondergrens en/of bovengrens
@attributes = grep { defined $_->{rangeLower} || defined $_->{rangeUpper}} @attributes;

# het LDAP-attribuut is Microsoft specifiek (voor Active Directory)
# Staat niet echt in de documentatie zie ik
@attributes = grep {substr( $_->{attributeID}, 0, 15 ) eq "1.2.840.113556."} @attributes;

@attributes = grep {$_->{systemOnly}} @attributes;

