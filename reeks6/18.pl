# Lol, count kunt ge niet gebruiken :p
# drie attributen : Count, Filter, Hints

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

# De klassenaam kan je bijvoorbeeld zien door naar de class kolom te kijken bij de users
@ARGV = ("user");

my $Users = bind_object("CN=Users,DC=iii,DC=hogent,DC=be");
$Users->{Filter} = shift;
print "$_->{adspath}\n" foreach in $Users;

print "\nAdministrator:\n";
my $a = $Users->GetObject("user","CN=Administrator");
print $a->{ADsPath};

