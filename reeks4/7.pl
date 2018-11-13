use Win32::OLE;

# Script stopt en geeft foutmelding als er iets niet werkt.
Win32::OLE->Option(Warn => 3);

# Informatie
$ComputerName = ".";
$NameSpace = "root/cimv2";
$ClassName = "Win32_NetworkAdapter";

# monik
$monik = "winmgmts://$ComputerName/$NameSpace";

#WbemServices laden
$WbemServices = Win32::OLE->GetObject($monik);

#Object laden
$NetworkAdapter = $WbemServices->get($ClassName);
$instances = $NetworkAdapter->Instances_();
printf("Aantal klassen: %d\n", $instances->Count);