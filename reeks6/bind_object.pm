
package bind_object;
use Win32::OLE 'in';
use Data::Dumper;
use Win32::OLE::Const "Active DS Type Library";  #bibliotheek inladen

sub get_constanten {
    %constanten = %{Win32::OLE::Const->Load("Active DS Type Library")};
    return %constanten;
}

sub get_omgekeerde_constanten {
    %constanten = get_constanten();
    while( ($key, $value) = each(%constanten) ){
        $omgekeerde_constanten{$value} = {$key};
    }
    return %omgekeerde_constanten;
}

sub bind_object {
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

1;