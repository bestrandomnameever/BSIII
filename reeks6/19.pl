# Die partities zijn dus gewoon die dingen waarmee je kan connecteren

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

my $RootDSE = bind_object("RootDSE");
Dumper $RootDSE->{dnsHostName};# Initialiseren, anders wordt niets geladen

# Kijk eens naar de RootDSE class in MSDN
# 2018-12-07-09-52-50.png

# Hier staan defaultNamingContext, configurationNamingContext en schemaNamingContext als attributen
# Je zou die dus zou kunnen uitschrijven
# Maar ik ga dat doen met de array namingContexts


foreach (@{$RootDSE->{namingContexts}}){
    print "$_\n"; # Je moet hier dus niet de distinguishedName ophalen, hmm
}