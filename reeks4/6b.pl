# Met Instances Of
use Win32::OLE 'in';

#Script stopt en geeft foutmelding als er iets niet werkt.
Win32::OLE->Option(Warn => 3);

#Informatie
$ComputerName = ".";
$NameSpace = "root/cimv2";
$ClassName = "Win32_OperatingSystem";

#Locator
$Locator = Win32::OLE->new("WbemScripting.SwbemLocator");

#Services
$WbemServices = $Locator->ConnectServer($ComputerName, $NameSpace);

#OperatingSystem laden
$OS = $WbemServices->get($ClassName);

#Nu een instance opvragen (er is maar 1)
$ObjectSet = $WbemServices->ExecQuery("SELECT * from $ClassName");

#Printen
( $instance ) = in $ObjectSet;

printf("objecttype van klasse = %s\n", Win32::OLE->QueryObjectType($instance));

foreach $property (in $instance->Properties_){
    if($property->{Name} =~ /Windows|version|caption|architecture/i){
        printf("%s => %s\n", $property->{Name}, $property->{value});  
    } 
}