use Win32::OLE::Const 'Active DS Type Library';
use Win32::OLE qw(in);
use Data::Dumper;
use Win32::OLE::Variant;

sub bind_object {
    my $parameter = shift;
    my $moniker = $parameter =~ /LDAP:////.*/ ? $parameter : "LDAP://193.190.126.71/$parameter";
    my $credentials = "Cedric Vanhaverbeke";

    my $dso = Win32::OLE->GetObject("LDAP:");
    my $Object = $dso->OpenDSObject($moniker, $credentials, $credentials, ADS_SECURE_AUTHENTICATION);
    return $Object;
}

sub print_error {
        print Win32::OLE->LastError() if Win32::OLE->LastError()
}

my $RootObj = bind_object("RootDSE");
print_error;

my $ADOConnection = Win32::OLE->new("ADODB.Connection");
$ADOConnection->{Provider} = "ADsDSOObject";
$ADOConnection->{Properties}{"User ID"} = "Cedric Vanhaverbeke";
$ADOConnection->{Properties}{"Password"} = "Cedric Vanhaverbeke";
$ADOConnection->{Properties}{"Encrypt Password"} = True;
$ADOConnection->Open();

my $ADOCommand = Win32::OLE->new("ADODB.Command");
$ADOCommand->{ActiveConnection} = $ADOConnection;
$ADOCommand->{Properties}{"Page Size"} = 20;

my $sBase = (bind_object($RootObj->Get("defaultNamingContext")))->{AdsPath};
my$sFilter = "(&(objectClass=user)(samAccountName=*))";
my $sAttributes = "samAccountName";
my $sScope      = "subtree";

$ADOCommand->{CommandText} = "<$sBase>;$sFilter;$sAttributes;$sScope";

my $ADOrecordset = $ADOCommand->Execute();
print_error;


print "Aantal objecten opgehaald " . $ADOrecordset->{RecordCount} . "\n";

$maxLength = -1;
while(!$ADOrecordset->{EOF}){
    my $samAccountName = $ADOrecordset->Fields("samAccountName")->{Value};
    $maxLength = length($samAccountName) if $maxLength < length($samAccountName);
    $ADOrecordset->MoveNext();
}

printf ("Langste lengte is %d\n", $maxLength);