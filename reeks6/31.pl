use bind_object;
use Win32::OLE 'in';
use Win32::OLE::Variant;
use Win32::OLE::Const "Active DS Type Library";
use Data::Dumper;


my $informaticaUsers = bind_object::bind_object("OU=EM7INF,OU=Studenten,OU=iii,DC=iii,DC=hogent,DC=be");
my $docenten = bind_object::bind_object("OU=Docenten,OU=iii,DC=iii,DC=hogent,DC=be");

my %constanten = %{Win32::OLE::Const->Load("Active DS Type Library")};

$number_to_type{$constanten{$_}} = $_ foreach(keys %constanten);

@argumentenlijst = ("Cedric Vanhaverbeke", "name");

if(scalar $ARGV != 2){
    print "Using default parameters:\n";
    print Dumper \@argumentenlijst;
    print "\n\n";
    @ARGV = @argumentenlijst;
}



# In de oplossing doen ze het gewoon met een get
# Dit was wellicht een slimmer idee :p
push @users, $_ foreach in $informaticaUsers;
push @users, $_ foreach in $docenten;
sub Zoeknaam {
        my $naam = shift;
        foreach(@users){
        # Getex geeft een referentie naar een array terug, hier is dat niet nodig
        # Want we willen gewoon kijken of de naam gelijk is
        # We vinden het attribuut name hiervoor
        if($_->Get("name") eq $ARGV[0]){
            printf("naam %s gevonden\n", $ARGV[0]);
            return $_;
        }
    }
    return 0;
}

use Math::BigInt;
sub convert_BigInt_string{
    my ($high,$low)=@_;
    my $HighPart = Math::BigInt->new($high);
    my $LowPartb = Math::BigInt->new($low);
    my $Radix    = Math::BigInt->new('0x100000000'); #dit is 2^32
    $LowPart+=$Radix if ($LowPart<0); #als unsigned int interperteren

    return ($HighPart * $Radix + $LowPart);
}

# We moeten deze stap doen met de LDAP interface.
sub zoekAttribuut {
    ( my $User, my $AttribuutNaam) = @_;
    if(isSet($User, $AttribuutNaam)){
        # Ga het nu gaan ophalen
        my $AttribuutWaarden = $User->Item($AttribuutNaam)->{Values};
        printWaarden($User->Item($AttribuutNaam), $AttribuutWaarden);
    } else {
        print "Het attribuut is niet ingesteld";
    }
    
}

sub isSet {
    (my $User, my $AttribuutNaam) = @_;
    $User->getInfoEx([$AttribuutNaam], 0); # Gaan ophalen, en moet zo omdat het geconstrueerd kan zijn

    # Returnt het hoeveelste element het attribuut is van de $User
    # Eerst een array maken van 0 tot $User->{propertyCount} -1
    # Daarna alles greppen dat dezelfde attribuutnaam heeft ams $Attribuutnaam
    # Aangezien er maar 1 attribuut zo kan zijn, neem het 0'de element
    return (grep {lc($User->Item($_)->{Name}) eq lc($AttribuutNaam)} 0..$User->{propertyCount}-1)[0];
}

#TODO: zoeken naar waarom er bij mij altijd een hash wordt geprint ipv een waarde
sub printWaarden {
    (my $Attribuut, my $AttribuutWaarden ) = @_;
    foreach(@{$AttribuutWaarden}){
        if ( $Attribuut->{ADsType} == ADSTYPE_NT_SECURITY_DESCRIPTOR) {
             $sec_object=$_->GetObjectProperty($Attribuut->{ADsType});
             $suffix = "eigenaar is ..." . $sec_object->{owner};
        }
        elsif ( $Attribuut->{ADsType} == ADSTYPE_OCTET_STRING) {
            $inhoud=$_->GetObjectProperty($Attribuut->{ADsType});
            $suffix = sprintf "%*v02X ","",$inhoud;
        }
        elsif ( $Attribuut->{ADsType} == ADSTYPE_LARGE_INTEGER ) {
            $inhoud=$_->GetObjectProperty($Attribuut->{ADsType});
            $suffix = (convert_BigInt_string($inhoud->{HighPart},$inhoud->{LowPart}));
        } 
        else  {
            $suffix = $_->GetObjectProperty($Attribuut->{ADsType});
            }
        printf("%-20s => %-20s\n", "waarde", $suffix);
    }
}

my $User = Zoeknaam($ARGV[0]);

if($User){
    zoekAttribuut($User, $ARGV[1]);
} else {
    print "naam niet gevonden\n";
}
