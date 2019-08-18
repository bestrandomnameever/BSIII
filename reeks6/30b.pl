use bind_object;
use Data::Dumper;
use Win32::OLE::Variant;

# Zeker zorgen dat de type library ook geladen is
my %constanten = %{Win32::OLE::Const->Load("Active DS Type Library")};

$number_to_type{$constanten{$_}} = $_ foreach(keys %constanten);

my $RootObj = bind_object::bind_object("RootDSE");
my $administrator = bind_object::bind_object("CN=Administrator,CN=Users,".$RootObj->get(defaultNamingContext));

# Getinfo doen om de attributen op te halen (niet de geconstrueerde dus)
$administrator->GetInfo();

# Property count is een attribuutEntry uit de IADsPropertyList interface
print "Aantal attributen in de Property Cache: $administrator->{PropertyCount}\n";

for(0 .. $administrator->{PropertyCount} - 1){
        my $attribuutEntry = $administrator->Next();
        # of my $attribuutEntry = $administrator->Item($i); #alternatief
        # Hoe the fuck kent ge de naam ADsType? In de documentatie staat er ADS_Type
        printf("%-35s (%d) : %-20s\n", $attribuutEntry->{Name}, $attribuutEntry->{ADsType} , getPropertyValue($attribuutEntry));
}

sub getPropertyValue {
    my $attribuutEntry = shift;
    my @Values = @{$attribuutEntry->{Values}};

    # In de oplossing lopen ze alle values van de entry af. Dit is omdat entries ook multivalued kunnen zijn
    foreach(@Values){
        if ( $attribuutEntry->{ADsType} == ADSTYPE_NT_SECURITY_DESCRIPTOR) {
             my $sec_object= $_->GetObjectProperty($attribuutEntry->{ADsType});
            $suffix = ("eigenaar is ..." . $sec_object->{owner});
        } elsif ( $attribuutEntry->{ADsType} == ADSTYPE_OCTET_STRING) {
            my $inhoud = $_->GetObjectProperty($attribuutEntry->{ADsType});
            $suffix = (sprintf "%*v02X ","",$inhoud);
        }   elsif ( $attribuutEntry->{ADsType} == ADSTYPE_LARGE_INTEGER ) {
            my $inhoud=($_->GetObjectProperty($attribuutEntry->{ADsType}));
            $suffix = (convert_BigInt_string($inhoud->{HighPart},$inhoud->{LowPart}));
        }   else  {
            $suffix = ($_->GetObjectProperty($attribuutEntry->{ADsType}));
        }
    }
    return $suffix;
}

#een groot geheel getal wordt teruggegeven als twee gehele getallen
#vb  29868835, -1066931206
# Het groot geheel getal dat hierbij hoort moet je berekenen als volgt :
# 29868835 . 2^32 + (2^32 - 1066931206 ) = 128285672722656250
# Onderstaande functie berekent deze waarde met behulp van de module Math::BigInt.
# Een bijkomend probleem is dat je deze waarde enkel met print juist uitschrijft,
# gebruik je printf dan moet je %s en niet %g gebruiken, anders krijg je een "afgeronde" waarde.
use Math::BigInt;
sub convert_BigInt_string{
    my ($high,$low)=@_;
    my $HighPart = Math::BigInt->new($high);
    my $LowPart = Math::BigInt->new($low);
    my $Radix = Math::BigInt->new('0x100000000'); #dit is 2^32
    $LowPart+=$Radix if ($LowPart<0); #als unsigned int interperteren

    return ($HighPart * $Radix + $LowPart);
}