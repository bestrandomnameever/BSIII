use Win32::OLE;
use bind_object;
use Data::Dumper;

@ARGV = (9041);

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

# Users zitten in studenten organizational unit
my $domein = bind_object::bind_object("OU=Studenten,OU=iii," . $RootDSE->{defaultNamingContext});

# Nu ldap query opstellen
my $sBase = "<$domein->{ADsPath}>";
my $sFilter = "(postalCode=$ARGV[0])"; # Postcode moet gelijk zijn aan het argument
my $sAttr = "name,postalcode";
my $sScope = "subtree";

$ADOcommand->{CommandText} = "$sBase;$sFilter;$sAttr;$sScope";

my $ADOResultset = $ADOcommand->Execute();
Win32::OLE->LastError()&& die (Win32::OLE->LastError());

while(!$ADOResultset->{EOF}){
    printf("%-20s\t", $ADOResultset->Fields($_)->{Value}) foreach split(",", $sAttr);
    print "\n";
    $ADOResultset->MoveNext();
}

# Alles closen
$ADOResultset->Close();
$ADOconnection->Close();