use Win32::OLE;

#Informatie
$ComputerName = ".";
$NameSpace = "root/cimv2";
$ClassName = "Win32_Directory";

#Locator
$Locator = Win32::OLE->new("WbemScripting.SwbemLocator");

#Services
$WbemServices = $Locator->ConnectServer($ComputerName, $NameSpace);

#Directory laden
$Directory = $WbemServices->get($ClassName);

#Associators bepalen
#Parameter moet op true staan: Schema only omdat we met klasses bezig zijn, geen instanties
$Associators = $Directory->Associators_(undef, undef, undef, undef, undef, 1);

printf("Er zijn %d associatorklassen", $Associators->{Count});
