use Win32::OLE;
use bind_object;
use Data::Dumper;

@ARGV = ("name");

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

my $sFilter = "($ARGV[0]=*)"; # Is er iets ingevuld?

my $sAttr = "$ARGV[0],objectClass";
my $sScope = "subtree";

$ADOcommand->{CommandText} = "$sBase;$sFilter;$sAttr;$sScope";

my $ADOResultset = $ADOcommand->Execute();
Win32::OLE->LastError()&& die (Win32::OLE->LastError());

while(!$ADOResultset->{EOF}){
    $objectclass=$ADOResultset->Fields("objectClass")->{Value};
    $class=$objectclass->[-1]; #laatste element
    $classes{$class}++;
    $ADOResultset->MoveNext();
}

$ADOResultset->Close();
$ADOconnection->Close();
foreach $cl (keys %classes){
   print $cl,"\n";
 }

#canonicalName is een geconstrueerd attribuut en mag niet in de filter worden opgenomen.
