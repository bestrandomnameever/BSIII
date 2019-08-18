use Win32::OLE;
use bind_object;
my $ADOconnection = Win32::OLE->new("ADODB.Connection");
   $ADOconnection->{Provider} = "ADsDSOObject";
   if ( uc($ENV{USERDOMAIN}) ne "III") { #niet ingelogd op het III domein
       $ADOconnection->{Properties}->{"User ID"}          = "Cedric Vanhaverbeke"; # vul in of zet in commentaar op school
       $ADOconnection->{Properties}->{"Password"}         = "Cedric Vanhaverbeke"; # vul in of zet in commentaar op school
       $ADOconnection->{Properties}->{"Encrypt Password"} = True;
   }

   $ADOconnection->Open();     
   
                                     # mag je niet vergeten
my $ADOcommand = Win32::OLE->new("ADODB.Command");
   $ADOcommand->{ActiveConnection}      = $ADOconnection;        # verwijst naar het voorgaand object
   $ADOcommand->{Properties}->{"Page Size"} = 20;

Win32::OLE->LastError()&& die (Win32::OLE->LastError());

my $RootDSE=bind_object::bind_object("RootDSE"); #serverless Binding
   $RootDSE->getinfo();
my $domein  = bind_object::bind_object( $RootDSE->Get("defaultNamingContext") );
my $sBase  = $domein->{adspath};

my $sFilter     = "(&(objectCategory=printQueue)(printColor=TRUE)(printDuplexSupported=TRUE))";
my $sAttributes = "printerName,printRate,printMaxResolutionSupported,printMaxResolutionSupported,printStaplingSupported";

my $sScope      = "subtree";
   $ADOcommand->{CommandText} = "<$sBase>;$sFilter;$sAttributes;$sScope";
   $ADOcommand->{Properties}->{"Sort On"} = "printerName";

my $ADOrecordset = $ADOcommand->Execute();
Win32::OLE->LastError()&& die (Win32::OLE->LastError());


print $ADOrecordset->{RecordCount}," AD-objecten\n";
print $ADOrecordset->Fields->{Count}," Ldap attributen opgehaald per AD-object\n";

use Win32::OLE::Variant;
use Win32::OLE qw(in);

while(! $ADOrecordset->{EOF}){
    @attributes = split(',', $sAttributes);
    foreach my $attribute (@attributes){
        printf("%s:%s;", $attribute, $ADOrecordset->Fields($attribute)->{Value});
    }
    print "\n" x 3;
    $ADOrecordset->MoveNext();
}


$ADOrecordset->Close();
$ADOconnection->Close();