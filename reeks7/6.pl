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

my $ADOcommand = Win32::OLE->new("ADODB.Command");
$ADOcommand->{ActiveConnection}      = $ADOconnection;        # verwijst naar het voorgaand object
$ADOcommand->{Properties}->{"Page Size"} = 20;     

Win32::OLE->LastError()&& die (Win32::OLE->LastError());

my $RootDSE=bind_object::bind_object("RootDSE"); #serverless Binding

$RootDSE->getinfo();

my $domein  = bind_object::bind_object( $RootDSE->Get("defaultNamingContext") );
my $sBase  = $domein->{adspath};

my $sFilter     = "(name=Piet*)";
my $sAttr       = "name";
my $sScope      = "subtree";

$ADOcommand->{CommandText} = "<$sBase>;$sFilter;$sAttr;$sScope";

my $ADORecordSet = $ADOcommand->Execute();

while(! $ADORecordSet->{EOF}){
    printf("%s\n", $ADORecordSet->Fields("$sAttr")->{Value});
    $ADORecordSet->MoveNext();
}


# Als ingevuld is met * gaat hij de cn printen