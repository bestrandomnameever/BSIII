use Win32::OLE;
use Data::Dumper;
use Win32::OLE::Const "Active DS Type Library";  #bibliotheek inladen

sub bind_object{
   my $parameter=shift;
   my $moniker;
   if ( uc($ENV{USERDOMAIN}) eq "III") { #ingelogd op het III domein
       $moniker = (uc( substr( $parameter, 0, 7) ) eq "LDAP://" ? "" : "LDAP://").$parameter;
       return (Win32::OLE->GetObject($moniker));
   }
   else {                                #niet ingelogd
       my $ip="193.190.126.71";          #voor de controle thuis
       $moniker = (uc( substr( $parameter, 0, 7) ) eq "LDAP://" ? "" : "LDAP://$ip/").$parameter;
       my $dso = Win32::OLE->GetObject("LDAP:");
       my $loginnaam = "Cedric Vanhaverbeke";          #vul in
       my $paswoord  = "Cedric Vanhaverbeke";          #vul in
       return ($dso->OpenDSObject( $moniker, $loginnaam, $paswoord, ADS_SECURE_AUTHENTICATION ));
  }
}

# Distinguished name
my $Name = "CN=BELIAL,OU=Domain Controllers,DC=iii,DC=hogent,DC=be";

$moniker_thuis="LDAP://193.190.126.71/$Name";
$obj=bind_object($moniker_thuis);
print Win32::OLE->LastError();

# Nu de attributen uitschrijven
# Die kan je vinden in MSDN in ADSI Objects of LDAP > Property >IADs

my @Properties = ("AdsPath", "Class", "GUID", "Name", "Parent", "Schema");
foreach(@Properties){
    printf("%s => %s\n", $_, $obj->$_);
}
