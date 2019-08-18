use Win32::OLE 'in';
use Data::Dumper;
use Win32::OLE::Const "Active DS Type Library";  #bibliotheek inladen

my %contextHash = (0 => "defaultNamingContext", 1 => "schemaNamingContext", 2 => "configurationNamingContext");

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

sub print_klassen {
    (my $class, my %classes) = @_;
    print "$class: \n";
    while( ($key, $value) = each(%classes) ){
        print " " x 5, "$key ==> $value\n";
    }
}



# Dit is recursief
sub stel_hash_op {
    my $object = bind_object(shift);
    %classes = ();
    if(in $object){
        foreach (in $object){
            $classesPerParent->{$object->{distinguishedName}}{$_}++;
        }
        print Dumper $classesPerParent;
        stel_hash_op($_->{distinguishedName}) foreach(keys %{$classesPerParent->{$object->{distinguishedName}}});
    }
}

sub stel_klassenaantal_op {
    %klassen = ();
    my $object = shift;
    foreach(keys %{$object}){
        $klassen{$_->{class}}++;
    }
    return %klassen;
}

sub print_aantal {
while( ($key, $value) = each (%{$classesPerParent})){
    print "Klassen in $key:\n";
    %klassen = stel_klassenaantal_op($value);
    while( ($key, $value) = each (%klassen) ){
        print " " x 5, "$key ==> $value\n";
    }
}
}

my $RootDSE = bind_object("RootDSE");
$RootDSE->{dnsHostName};# Initialiseren, anders wordt niets geladen
@ARGV= ( 0 );

my $context = $RootDSE->{ $contextHash{$ARGV[0]} };
stel_hash_op($context);
print_aantal();


