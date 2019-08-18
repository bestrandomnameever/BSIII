# De klasse Win32_OperatingSystem heeft die info
use Win32::OLE 'in';
use Data::Dumper;

#Script stopt en geeft foutmelding als er iets niet werkt.
Win32::OLE->Option(Warn => 3);

$ComputerName = ".";
$NameSpace = "root/cimv2";

$ClassName = "Win32_OperatingSystem=@";

$moniker = "winmgmts://$ComputerName/$NameSpace:$ClassName";
$OperatingSystem = Win32::OLE->GetObject($moniker);

print Dumper $OperatingSystem;


foreach $property (in $OperatingSystem->Properties_){
    if($property->{Name} =~ /Windows|version|caption|architecture/i){
        printf("%s => %s\n", $property->{Name}, $property->{value});  
    } 
}