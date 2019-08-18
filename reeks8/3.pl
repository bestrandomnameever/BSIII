use Win32::OLE qw(in);
use Win32::OLE::Const 'Active DS Type Library';
use Win32::OLE::Variant;
use Data::Dumper;

@ARGV = ("group");

sub bind_object {
    my $parameter = shift;
    my $moniker = uc($parameter) =~ /LDAP:////.*/ ? $parameter : "LDAP://193.190.126.71/$parameter";
    my $credentials = "Cedric Vanhaverbeke";

    my $dso = Win32::OLE->GetObject("LDAP:");
    my $Object = $dso->OpenDSObject($moniker, $credentials, $credentials, ADS_SECURE_AUTHENTICATION);
    return $Object;
}

sub print_error {
        print Win32::OLE->LastError() if Win32::OLE->LastError()
}

my $RootObj = bind_object("RootDSE");
my $AbstractClass = bind_object("schema/$ARGV[0]");
print_error;

@MandatoryProperties = @{$AbstractClass->{MandatoryProperties}};
my $attribuut = "cn";

my $ADOConnection = Win32::OLE->new("ADODB.Connection");
$ADOConnection->{Provider} = "ADsDSOObject";
$ADOConnection->{Properties}{"User ID"} = "Cedric Vanhaverbeke";
$ADOConnection->{Properties}{"Password"} = "Cedric Vanhaverbeke";
$ADOConnection->{Properties}{"Encrypt Password"} = True;
$ADOConnection->Open();
print_error;

my $ADOCommand = Win32::OLE->new("ADODB.Command");
$ADOCommand->{ActiveConnection} = $ADOConnection;
$ADOCommand->{Properties}{"Page size"} = 20;

my $sBase = bind_object($RootObj->Get("defaultNamingContext"))->{ADsPath};
my $sFilter = "(objectClass=$ARGV[0])";
my $sAttr = "cn";
my $sScope = "subtree";

$ADOCommand->{CommandText} = "<$sBase>;$sFilter;$sAttr;$sScope";

my $ResultSet = $ADOCommand->Execute();
printf("Er zijn in het totaal %d klassen", $ResultSet->{RecordCount});

while(!$ResultSet->{EOF}){
    printf("cn = %s\n", $ResultSet->Fields("cn")->{Value});
    $ResultSet->MoveNext();
}

$ResultSet->Close();
$ADOConnection->Close();

print_error;