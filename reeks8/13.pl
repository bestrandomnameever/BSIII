use Win32::OLE 'in';
use Win32::OLE::Const 'Active DS Type Library';
use Win32::OLE::Variant;
use bind_object;
use Data::Dumper;

my $RootDSE = bind_object::bind_object("RootDSE");

# Connection
my $ADOConnection = Win32::OLE->new("ADODB.Connection");
$ADOConnection->{Provider} = "ADsDSOObject";
if(uc($ENV{USERDOMAIN} ne 'III')){
    $ADOConnection->{Properties}{"User ID"} = "Interim F";
    $ADOConnection->{Properties}{"Password"} = "Interim F";
    $ADOConnection->{Properties}{"Encrypt Password"} = True;
}

$ADOConnection->Open();

# Command
my $ADOCommand = Win32::OLE->new("ADODB.Command");
$ADOCommand->{ActiveConnection} = $ADOConnection;
$ADoCommand->{Properties}{"Page Size"} = 20;

# Simuleren van de organizational unit
my $domein = bind_object::bind_object("OU=EM7INF,OU=Studenten,OU=iii," . $RootDSE->Get("defaultNamingContext"))->{AdsPath};

my $sBase = $domein;
my $sFilter = "(&(objectCategory=person)(objectClass=user))";
my $sAttributes = "distinguishedName";
my $sScope = "subtree";

$ADOCommand->{CommandText} = "<$sBase>;$sFilter;$sAttributes;$sScope";
my $ADOrecordset = $ADOCommand->Execute();


while(!$ADOrecordset->{EOF}){
    my $dn = $ADOrecordset->Fields("distinguishedName")->{Value};
    my $User = bind_object::bind_object($dn);
    #$container->MoveHere($User->{AdsPath});
    die "Simulatie van verplaatsen" if Win32::OLE->LastError();
    $ADOrecordset->MoveNext();
}
