#. . . implementatie functie bind_object zie sectie 5 in reeks 6
#. . . implementatie functie valueattribuut: zie oefening 1

use Win32::OLE::Const "Active DS Type Library"; #inladen constanten voor PutEx

my $RootDSE=bind_object("RootDSE"); #serverless Binding
   $RootDSE->getinfo();

my $cont=bind_object("ou=...,ou=...,ou=labo,".$RootDSE->{defaultNamingContext}); #vul aan
my $groep = bind_object("cn=...,ou=...,ou=...,ou=labo,".$RootDSE->{defaultNamingContext});  #vul aan

#alle users van de container ophalen
  $cont->{filter}=[user];

  foreach (in $cont) {
    push @leden ,$_->{distinguishedName};
  }

  $groep->PutEx(ADS_PROPERTY_UPDATE,"member",\@leden); #ADS_PROPERTY_UPDATE=2
  $groep->SetInfo();

#controleer een user
my $user=bind_object("cn=...,ou=...,ou=...,ou=labo,".$RootDSE->{defaultNamingContext});
print join(",",@{valueattribuut($user,"memberOf")});