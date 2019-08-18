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

my $domein = bind_object::bind_object("OU=PC's,OU=iii," . $RootDSE->{defaultNamingContext});
my $sBase = "<$domein->{ADsPath}>";

my $sFilter = "(name=*A)"; # Je mag hier dus nooit een geconstrueerd attribuut zetten. Stout!
my $sAttr = "name";
my $sScope = "subtree";

$ADOcommand->{CommandText} = "$sBase;$sFilter;$sAttr;$sScope";

$ADOResultSet = $ADOcommand->Execute();
Win32::OLE->LastError()&& die (Win32::OLE->LastError()); # Is er een fout gebeurtd?

while(!$ADOResultSet->{EOF}){
    printf("%s\n", $ADOResultSet->Fields($sAttr)->{Value});
    $ADOResultSet->MoveNext();
}