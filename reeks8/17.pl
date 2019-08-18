use Win32::OLE qw(in);
use Win32::OLE::Const 'Active DS Type Library';
use Win32::OLE::Variant;
use Data::Dumper;

my %provincies = (
    1 => "Brabant",
    2 => "Antwerpen",
    3 => "Limburg", 4 => "Luik", 5 => "Namen", 6 => "Luxemburg", 7 => "Henegouwen" , 8 => "WVlaanderen", 9 => "OVlaanderen"
);

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

sub voeg_groep_toe {
    
}

my $RootObj = bind_object("RootDSE");
$realcontainer = bind_object("OU=EM7INF,OU=Studenten,OU=iii," . $RootObj->Get("defaultNamingContext"));
$realcontainer->{Filter} = ["user"];



@ARGV = "stdinwe.out";
while(<>){
     chomp;
     @lijn = split(":", $_);
     my $cn = $lijn[1];
     $lijn[-1] =~ /([a-zA-Z]*)([0-9]*)/;
     $richting = $1;
     $nummer = $2;
     $richingen{$richting}++;

     $gevondenProvincies{$provincies{substr($lijn[4], 0, 1)}}++;
}

print Dumper \%richingen;
print Dumper \%gevondenProvincies;