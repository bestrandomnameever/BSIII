use bind_object;
use Data::Dumper;
use Win32::OLE::Variant;

my $RootDSE = bind_object::bind_object("RootDSE");
my $defaultNamingContext = $RootDSE->get(defaultNamingContext);
my $schemaNamingContext = $RootDSE->get(schemaNamingContext);

my $Schema = bind_object::bind_object($schemaNamingContext);
$schemaHash{$_->{Class}}++ foreach (in $Schema);

while( (my $key, my $value) = each(%schemaHash) ){
    printf("%-20s => %-20d\n", $key, $value);
}