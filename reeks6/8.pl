use Win32::OLE 'in';
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

my $distinguishedName = "CN=Cedric Vanhaverbeke,OU=EM7INF,OU=Studenten,OU=iii,DC=iii,DC=hogent,DC=be";
my $obj = bind_object($distinguishedName);

# ADSI Attributen
print $obj->GUID, "\n";
print $obj->ADsPath, "\n";
print $obj->Class, "\n";

print "\n";
# LDAP Attributen
print sprintf ("%*v02X ","", $obj->objectGUID), "\n";
print $obj->distinguishedName, "\n";

print join ", ", @{$obj->objectClass}, "\n";
# Of enkel de juiste klasse uitschrijven
print $obj->objectClass->[-1], "\n";
