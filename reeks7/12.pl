#dsquery * "cn=configuration,dc=iii,dc=hogent,dc=be" -filter "(fromServer=*)" -attr objectclass
#onderstaande query geeft geen resultaten
#dsquery * "cn=configuration,dc=iii,dc=hogent,dc=be" -filter "(fromServer=*be)" -attr objectclass

# Om configuration eerst te zien in ADSI edit moet je daarmee eerst connecteren natuurlijk :)
use Win32::OLE;
use bind_object;
use Data::Dumper;

@ARGV = ("objectclass");

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

my $domein = bind_object::bind_object($RootDSE->{configurationNamingContext});
my $sBase = "<$domein->{ADsPath}>";

my $sFilter = "($ARGV[0]=*)"; # Is er iets ingevuld?

my $sAttr = "$ARGV[0],objectClass";
my $sScope = "subtree";

$ADOcommand->{CommandText} = "$sBase;$sFilter;$sAttr;$sScope";

my $ADOResultset = $ADOcommand->Execute();
Win32::OLE->LastError()&& die (Win32::OLE->LastError());

while(!$ADOResultset->{EOF}){
    printf("%-20s\t", $ADOResultset->Fields($_)->{Value}[-1]) foreach split(',', $sAttr);
    print "\n";
    $ADOResultset->MoveNext();
}

$ADOResultset->Close();
$ADOconnection->Close();
foreach $cl (keys %classes){
   print $cl,"\n";
 }

