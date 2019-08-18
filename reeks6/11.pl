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

sub GetDomeinDN {
    my $RootObj = bind_object("RootDSE");

    # Eerst dnsHostName opvragen voor je iets anders kan opvragen. De geheimen van LDAP
    $RootObj->{dnsHostName};
    return $RootObj->{defaultNamingContext};
}

my $DomeinDN = GetDomeinDN();

my $distinguishedName = "CN=Cedric Vanhaverbeke,OU=EM7INF,OU=Studenten,OU=iii," . $DomeinDN;
my $obj = bind_object($distinguishedName);
print Dumper $obj;