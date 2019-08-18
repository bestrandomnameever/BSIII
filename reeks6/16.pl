# dsquery  computer -s 193.190.126.71 -u "Cedric Vanhaverbeke" -p "Cedric Vanhaverbeke"  "dc=iii,dc=hogent,dc=be"
# Je krijgt iets in de aard van "CN=MONTEVERDI,OU=225,OU=PC's,OU=iii,DC=iii,DC=hogent,DC=be"
# Hier zie je dat je moet zoeken naar de OU PC's. Het lokaal is blijkbaar 225

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

my $Classroom = bind_object("OU=225,OU=PC's,OU=iii,DC=iii,DC=hogent,DC=be");
# Geen idee waarom ze hier adspath printen. Dit is het LDAP path precies
print $Classroom->{adspath};
foreach(in $Classroom){
    print $_->{cn}, "\n";
}

