use Win32::OLE qw(in);
use Win32::OLE::Const 'Active DS Type Library';
use Win32::OLE::Variant;
use Data::Dumper;

sub bind_object {
    my $parameter = shift;
    my $moniker = $parameter =~ /LDAP:////.*/ ? $parameter : "LDAP://193.190.126.71/$parameter";
    my $credentials = "Interim F";

    my $dso = Win32::OLE->GetObject("LDAP:");
    my $Object = $dso->OpenDSObject($moniker, $credentials, $credentials, ADS_SECURE_AUTHENTICATION);
    return $Object;
}

sub print_error {
        print Win32::OLE->LastError() if Win32::OLE->LastError()
}

my $RootObj = bind_object("RootDSE");
my $Joris = bind_object("CN=jomoreau,OU=Docenten,OU=iii," . $RootObj->Get("defaultNamingContext"));

$Joris->GetInfoEx(["memberOf", "primaryGroupID"], 0);
# Gewone groepen
print Dumper $Joris->Get("memberOf"), $Joris->Get("primaryGroupID");




print_error;