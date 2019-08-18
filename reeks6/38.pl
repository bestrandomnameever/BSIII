use bind_object;
use Data::Dumper;
use Win32::OLE::Variant;

my $RootDSE = bind_object::bind_object("RootDSE");
$Schema = bind_object::bind_object($RootDSE->get(schemaNamingContext));

# Het abstract schema binden doe je met een speciaal ADsPath!
# LDAP://[server|domein/]schema
# Bij ons is dat gewoon hetzelfde als bind_object("schema");

my $AbsctractSchema = bind_object::bind_object("schema");
foreach(in $AbsctractSchema){
    $abstract{$_->{Class}}++;
}

print Dumper \%abstract;