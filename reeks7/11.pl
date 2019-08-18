#dsquery * "DC=iii,DC=hogent,DC=be"  -s 193.190.126.71 -u "Interim F" -p "Interim F"  -filter  "(cn=Piet*)" -attr cn primaryGroupID
#Met de tweede dsquery kan je de naam van alle primaire groepen terugvinden:
#dsquery * "DC=iii,DC=hogent,DC=be"  -s 193.190.126.71 -u "Interim F" -p "Interim F"  -filter  "(objectclass=group)" -attr cn primaryGrouptoken
#Ga na dat je de filter niet kan verfijnen zodat je enkel de gewenste groep ziet - lees ook de foutmelding die je krijgt hiervoor.

# Je kan primaryGroupID en primaryGrouptoken niet gaan verfijnen omdat het constructed attributen zijn
# Ze willen precies wel echt dat je daar goed op let :p

use Win32::OLE;
use bind_object;
use Data::Dumper;

@ARGV = ("Cedric Vanhaverbeke");


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

# Eerste command opbouwen

my $sBase = "<$domein->{ADsPath}>";

my $sFilter = "(cn=$ARGV[0])"; # Je mag hier dus nooit een geconstrueerd attribuut zetten. Stout!
my $sAttr = "cn,primaryGroupID";
my $sScope = "subtree";

$ADOcommand->{CommandText} = "$sBase;$sFilter;$sAttr;$sScope";

$ADOResultSet = $ADOcommand->Execute();
Win32::OLE->LastError()&& die (Win32::OLE->LastError()); # Is er een fout gebeurtd?

# Er kan maar 1 resultaat zijn dus ik ga yolo en steek dat niet in een lus waddup
$User{naam} = $ARGV[0];
$User{primaryGroupID} = $ADOResultSet->Fields(primaryGroupID)->{Value};

# Closen en dan een nieuw beginnen
$ADOResultSet->Close();
$ADOcommand->Close();

# Tweede command opbouwen

$ADOcommand = Win32::OLE->new("ADODB.Command");
$ADOcommand->{ActiveConnection} = $ADOconnection;        # verwijst naar het voorgaand object
$ADOcommand->{Properties}->{"Page Size"} = 20;     

Win32::OLE->LastError()&& die (Win32::OLE->LastError()); # Is er een fout gebeurtd?

$sFilter = "(objectclass=group)"; # Je mag hier dus nooit een geconstrueerd attribuut zetten. Stout!
$sAttr = "cn,primaryGrouptoken";
$sScope = "subtree";

$ADOcommand->{CommandText} = "$sBase;$sFilter;$sAttr;$sScope";
$ADOResultSet = $ADOcommand->Execute();
Win32::OLE->LastError()&& die (Win32::OLE->LastError()); # Is er een fout gebeurtd?

while(!$ADOResultSet->{EOF}){
    last if $ADOResultSet->Fields("primaryGrouptoken")->{Value} eq $User{primaryGroupID};
    $ADOResultSet->MoveNext();
}

$User{GroupName} = $ADOResultSet->Fields("cn")->{Value};
while( (my $key, my $value) = each(%User) ){
    printf("%-20s => %-20s\n", $key, $value);
}
$ADOResultSet->Close();
$ADOcommand->Close();






