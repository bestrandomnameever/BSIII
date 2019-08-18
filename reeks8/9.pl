use Win32::OLE 'in';
use Win32::OLE::Variant;
use bind_object;
use Data::Dumper;
use Win32::OLE::Const 'Active DS Type Library';

my $ADOconnection = Win32::OLE->new("ADODB.Connection");
   $ADOconnection->{Provider} = "ADsDSOObject";
   if ( uc($ENV{USERDOMAIN}) ne "III") { #niet ingelogd op het III domein
      $ADOconnection->{Properties}->{"User ID"}          = "Cedric Vanhaverbeke"; # vul in of zet in commentaar
      $ADOconnection->{Properties}->{"Password"}         = "Cedric Vanhaverbeke"; # vul in of zet in commentaar
      $ADOconnection->{Properties}->{"Encrypt Password"} = True;
   }
   $ADOconnection->Open();                                       # mag je niet vergeten
my $ADOcommand = Win32::OLE->new("ADODB.Command");
   $ADOcommand->{ActiveConnection}      = $ADOconnection;        # verwijst naar het voorgaand object
   $ADOcommand->{Properties}->{"Page Size"} = 20;

my $RootDSE=bind_object::bind_object("RootDSE"); #serverless Binding
   $RootDSE->getinfo();


my $domein = bind_object::bind_object( $RootDSE->{defaultNamingContext});

my $sBase = $domein->{adspath};
my $sFilter     = "(objectcategory=group)";
my $sAttributes = "cn,groupType";
my $sScope      = "subtree";

   $ADOcommand->{Properties}->{"Sort On"} = "groupType";
   $ADOcommand->{CommandText} = "<$sBase>;$sFilter;$sAttributes;$sScope";
my $ADOrecordset = $ADOcommand->Execute();

until ( $ADOrecordset->{EOF} ) {
    printf "%04b\t%s\n",$ADOrecordset->Fields("groupType")->{Value}   # eerste bit staat op 1
                       ,$ADOrecordset->Fields("cn")->{Value};
    $ADOrecordset->MoveNext();
}
$ADOrecordset->Close();
$ADOconnection->Close();
