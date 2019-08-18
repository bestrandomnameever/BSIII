use bind_object;
use Data::Dumper;
use Win32::OLE::Variant;

%constanten = %{ Win32::OLE::Const->Load("Active DS Type Library") };
$gefilterde{$_} = $constanten{$_} foreach grep {/.*systemflag.*/i} keys %constanten;
print Dumper \%gefilterde;

my $RootDSE = bind_object::bind_object("RootDSE");
my $schemaNamingContext = $RootDSE->get(schemaNamingContext);

my $Schema = bind_object::bind_object($schemaNamingContext);

$Schema->{Filter} = "attributeSchema";

push @attributes, $_ foreach (in $Schema);

@geconstrueerd = grep {$_->{systemFlag} & $gefilterde{ADS_SYSTEMFLAG_ATTR_IS_CONSTRUCTED}} @attributes;
@nietWijzigbaarDoorGebruiker = grep {$_->{systemOnly}} @attributes;

print "GECONSTRUEERDE:\n";
print "-" x 50, "\n"; 
print join ", ", map {$_->{ldapDisplayName}} @geconstrueerd;

print "NIET WIJZIGBARE:\n";
print "-" x 50, "\n"; 
print join ", ", map {$_->{ldapDisplayName}} @nietWijzigbaarDoorGebruiker;