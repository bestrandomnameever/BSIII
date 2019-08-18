# Als je gaat kijken in de AD Search Library zoals ze dat noemen
# de tak Directory Access Technologies / Active Directory Service Interfaces / Using Active Directory Service Interfaces / Searching Active Directory,

# Vind je memberof
# Je hebt blijkbaar ook member

use bind_object;
use Data::Dumper;
use Win32::OLE::Variant;

my $ADOconnection = Win32::OLE->new("ADODB.Connection");
$ADOconnection->{Provider} = "ADsDSOObject";


if ( uc($ENV{USERDOMAIN}) ne "III") { 
    $ADOconnection->{Properties}->{"User ID"}          = "Interim F";
    $ADOconnection->{Properties}->{"Password"}         = "Interim F";
    $ADOconnection->{Properties}->{"Encrypt Password"} = True;
}

$ADOconnection->Open();      

my $ADOcommand = Win32::OLE->new("ADODB.Command");
$ADOcommand->{ActiveConnection} = $ADOconnection;   
$ADOcommand->{Properties}->{"Page Size"} = 20;

my $RootDSE=bind_object::bind_object("RootDSE");
$RootDSE->getinfo();


my $domein = bind_object::bind_object( $RootDSE->Get("defaultNamingContext") );
my $sBase = $domein->{adspath};



my $sFilter = "(&(objectclass=group)(cn=beveiliging))";
my $sAttributes = "distinguishedName"; 
my $sScope      = "subtree";

$ADOcommand->{CommandText} = "<$sBase>;$sFilter;$sAttributes;$sScope";

my $ADOrecordset = $ADOcommand->Execute();

($ADOrecordset->{RecordCount} ==1) or die "$groep bestaat niet\n";

$distinguishedName =$ADOrecordset->Fields("distinguishedName")->{Value};

$ADOrecordset->Close();
$ADOcommand->Close();

$ADOcommand = Win32::OLE->new("ADODB.Command");
$ADOcommand->{ActiveConnection} = $ADOconnection;   
$ADOcommand->{Properties}->{"Page Size"} = 20;



my $sFilter = "(memberof:1.2.840.113556.1.4.1941:=$distinguishedName)"; # Let op van de haken in MSDN documentatie (da's dus niet juist)
my $sAttributes = "cn,distinguishedname"; 
my $sScope      = "subtree";

$ADOcommand->{CommandText} = "<$sBase>;$sFilter;$sAttributes;$sScope";

$ADOrecordset = $ADOcommand->Execute();

while(!$ADOrecordset->{EOF}){
    printf("%-20s\t", $ADOrecordset->Fields($_)->{Value}) foreach split(",", $sAttributes);
    print "\n";
    $ADOrecordset->MoveNext();
}

# Alle groepen van de eerste user hier uit de lijst
$ADOrecordset->MoveFirst();
my $firstPerson = $ADOrecordset->Fields("distinguishedname")->{Value};

$ADOrecordset->Close();

my $sFilter = "(member:1.2.840.113556.1.4.1941:=$firstPerson)";
my $sAttributes = "cn,distinguishedname"; 
my $sScope      = "subtree";

$ADOcommand->{CommandText} = "<$sBase>;$sFilter;$sAttributes;$sScope";

print "\n" x 5;

$ADOrecordset = $ADOcommand->Execute();

while(!$ADOrecordset->{EOF}){
    printf("%-20s\t", $ADOrecordset->Fields($_)->{Value}) foreach split(",", $sAttributes);
    print "\n";
    $ADOrecordset->MoveNext();
}






