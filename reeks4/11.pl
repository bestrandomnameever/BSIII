use Win32::OLE 'in';
use Data::Dumper;
#Simuleren van argumenten
@ARGV = (19);
@ARGV or die "Geef IRQ-nummer van de gewenste Interrupt Request";

($nummer) = @ARGV;

#Informatie
$ComputerName = "localhost";
$NameSpace = "root/cimv2";
$ClassName = "Win32_IRQResource";

#Wbemservices aanmaken
$monik = "winmgmts://$ComputerName/$NameSpace";
$Wbemservices = Win32::OLE->GetObject($monik);

#Instantie die voldoet aan een bepaalde voorwaarde opvragen, namelijk het nummer
$ObjectSet = $Wbemservices->ExecQuery("select * from $ClassName where IRQNumber=$nummer");
( $IRQ ) = (in $ObjectSet);

#Nu alle associaties van deze instantie met netwerk in de naam
$ObjectSet = $IRQ->Associators_();
@instances = (in $ObjectSet);
for (@instances){
    print $_->{Description}, "\n" if $_->__CLASS == "Win32_NetworkAdapter";
}

# Of beter:
#(my $Class) = in $WbemServices->ExecQuery("select * from Win32_IRQResource where IRQNumber = " . shift);
# Check de parameters in Associators_() in MSDN
#my $Associators = $Class->Associators_(undef, "Win32_NetworkAdapter");