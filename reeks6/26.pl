use bind_object;
use Data::Dumper;

# Variant gaat tijdstippe automatisch omzetten voor jou :D
use Win32::OLE::Variant;

my %constanten = (
    #Opletten wat je tussen de haken zet hier :
    E_ADS_PROPERTY_NOT_FOUND => Win32::OLE::HRESULT(0x8000500D)
);

$attribuut = @ARGV[0] ?  $ARGV[0] : "distinguishedName";
my $object = bind_object::bind_object("CN=Administrator,CN=Users,DC=iii,DC=hogent,DC=be");


# GetEx returnt een array, daarom $arrayref
my $arrayref = $object->GetEx($attribuut);

if(Win32::OLE->LastError() == $constanten{E_ADS_PROPERTY_NOT_FOUND}){
    # Eerst de info nog eens proberen ophalen als het attribuut geconstrueerd is
    # Omdat GetEx en Get niet zomaar werken bij geconstrueerde attributen
    $object->GetInfoEx([$attribuut], 0); # Met GetInfoEx moet je een array meegeven en eindigen met 0
    $arrayref = $object->GetEx($attribuut);
}

# Nu kijken of het nog altijd niet ingevuld is
if( Win32::OLE->LastError() == $constanten{E_ADS_PROPERTY_NOT_FOUND} ){
    print $attribuut, " is niet ingesteld";
} else {
    print "De waarde voor het attribuut is: \n";
    printf("%s\n", $_) foreach @{$arrayref};
}