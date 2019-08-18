use Win32::OLE;
use bind_object;
use Data::Dumper;

my $ADOconnection = Win32::OLE->new("ADODB.Connection");
$ADOconnection->{Provider} = "ADsDSOObject";

if ( uc($ENV{USERDOMAIN}) ne "III") { #niet ingelogd op het III domein
    $ADOconnection->{Properties}->{"User ID"}          = "Cedric Vanhaverbeke"; # vul in of zet in commentaar op school
    $ADOconnection->{Properties}->{"Password"}         = "Cedric Vanhaverbeke"; # vul in of zet in commentaar op school
    $ADOconnection->{Properties}->{"Encrypt Password"} = True;
}

$ADOconnection->Open();


my $ADOcommand = Win32::OLE->new("ADODB.Command");
$ADOcommand->{ActiveConnection} = $ADOconnection;        # verwijst naar het voorgaand object
$ADOcommand->{Properties}->{"Page Size"} = 20;     

Win32::OLE->LastError()&& die (Win32::OLE->LastError()); # Is er een fout gebeurtd?

my $RootDSE = bind_object::bind_object("RootDSE");

$RootDSE->GetInfo(); # Alle attribten ophalen van RootDSE

my $domein = bind_object::bind_object($RootDSE->{defaultNamingContext});
my $sBase = "<$domein->{ADsPath}>";

my $sFilter = ""; # Mag leeg zijn
#my $sFilter    = "(objectClass=user)";  #enkel users
#my $sFilter    = "(objectClass=u*)";    #lukt niet met wildcards
#my $sFilter    = "((distinguishedName=*) #lukt wel
#my $sFilter    = "((distinguishedName=CN=*)) #lukt niet meer
#my $sFilter    = "(cn=*an*)";            #lukt wel met wildcards
#my $sFilter    = "(cn=*)";               #attribuut cn moet ingesteld zijn

my $sAttr = "*";
my $sScope = "subtree";

$ADOcommand->{CommandText} = "$sBase;$sFilter;$sAttr;$sScope";

my $ADOResultset = $ADOcommand->Execute();
Win32::OLE->LastError()&& die (Win32::OLE->LastError());

print "Aantal objecten opgehaald: $ADOResultset->{RecordCount}\n";

$ADOResultset->Close();
$ADOconnection->Close();

