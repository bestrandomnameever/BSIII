use Win32::OLE 'in';
use Win32::OLE::Const 'Active DS Type Library';
use Data::Dumper;

# Aanmaken van dso
my $dso = Win32::OLE->GetObject("LDAP:");
my $Username = "Cedric Vanhaverbeke";
my $Password = "Cedric Vanhaverbeke";


my $Moniker = "LDAP://193.190.126.71/CN=BELIAL,OU=Domain Controllers,DC=iii,DC=hogent,DC=be";

# Deze laatste is een constante die je inlaadt via de Type Library
# Je kan een overzicht van de constanten vinden in MSDN bij de methode OpenDSObject (in de opgave staat hoe je eraan komt)
my $Object = $dso->OpenDSObject($Moniker, $Username, $Password, ADS_SECURE_AUTHENTICATION);