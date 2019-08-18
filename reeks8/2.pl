# We  kunnen maar 5 objecten aanmaken in de groep
# account, contact, organizationalUnit, user

use bind_object;
use Data::Dumper;
use Win32::OLE::Variant;
use Win32::OLE::Const "Active DS Type Library";

use bind_object;
use Data::Dumper;
use Win32::OLE::Variant;

my $ADOconnection = Win32::OLE->new("ADODB.Connection");
$ADOconnection->{Provider} = "ADsDSOObject";

if ( uc($ENV{USERDOMAIN}) ne "III") { 
$ADOconnection->{Properties}->{"User ID"} = "Interim F";
$ADOconnection->{Properties}->{"Password"} = "Interim F";
$ADOconnection->{Properties}->{"Encrypt Password"} = True;
}

$ADOconnection->Open();      

my $ADOcommand = Win32::OLE->new("ADODB.Command");
$ADOcommand->{ActiveConnection} = $ADOconnection;   
$ADOcommand->{Properties}->{"Page Size"} = 20;

my $RootDSE=bind_object::bind_object("RootDSE");
$RootDSE->getinfo();



my $domein = bind_object::bind_object("OU=Studenten,OU=iii,DC=iii,DC=hogent,DC=be");
my $sBase = $domein->{adspath};


my $sFilter = "(&(objectCategory=person)(description=0*/12/*))";
my $sAttributes = "description,distinguishedname"; 
my $sScope      = "subtree";

$ADOcommand->{CommandText} = "<$sBase>;$sFilter;$sAttributes;$sScope";

my $ADOrecordset = $ADOcommand->Execute();
Win32::OLE->LastError()&& die (Win32::OLE->LastError());

print "Er zijn $ADOrecordset->{RecordCount} users in de groep\n";

while(!$ADOrecordset->{EOF}){
    push @users , $ADOrecordset->Fields("distinguishedname")->{Value};
    $ADOrecordset->MoveNext();
}

$ADOrecordset->Close();
$ADOcommand->Close();

my $group = bind_object::bind_object("CN=december-0,OU=Cedric Vanhaverbeke,OU=Labo,DC=iii,DC=hogent,DC=be");



# Ik had een probleem, ik had de Active DS Library niet ingeladen
$group->PutEx(ADS_PROPERTY_UPDATE, "member", \@users);
$group->SetInfo();