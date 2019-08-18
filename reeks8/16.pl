use Win32::OLE qw(in);
use Win32::OLE::Const 'Active DS Type Library';
use Win32::OLE::Variant;
use Data::Dumper;

my $RootObj = bind_object("RootDSE");
$realcontainer = bind_object("OU=EM7INF,OU=Studenten,OU=iii," . $RootObj->Get("defaultNamingContext"));
$realcontainer->{Filter} = ["user"];

    foreach(in $realcontainer){
        $_->GetInfoEx(["canonicalName"], 0);
        if($_->Get("canonicalName")){
            @names = (split("/", $_->Get("canonicalName")));
            push @User, $names[-1];
        }
    }

sub bind_object {
    my $parameter = shift;
    my $moniker = $parameter =~ /LDAP:////.*/ ? $parameter : "LDAP://193.190.126.71/$parameter";
    my $credentials = "Cedric Vanhaverbeke";

    my $dso = Win32::OLE->GetObject("LDAP:");
    my $Object = $dso->OpenDSObject($moniker, $credentials, $credentials, ADS_SECURE_AUTHENTICATION);
    return $Object;
}

sub print_error {
        print Win32::OLE->LastError() if Win32::OLE->LastError()
}

sub geef_keuzes {

}

sub geef_overzicht {
    print Dumper \@User;
}

geef_overzicht;


while(choice != -1){
    
    
}