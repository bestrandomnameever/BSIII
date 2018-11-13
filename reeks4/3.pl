use Win32::OLE;
Win32::OLE->Option(Warn => 3);

$ComputerName = ".";
$NameSpace = "root/cimv2";

#Om een instantie op te vragen
#Mocht je de klasse willen dan was Win32_NetworkAdapter genoeg
$ClassName = "Win32_NetworkAdapter.DeviceID=\"0\"";

$moniker = "winmgmts://$ComputerName/$NameSpace:$ClassName";

$NetworkAdapter = Win32::OLE->GetObject($moniker);
printf("objecttype van klasse = %s\n", Win32::OLE->QueryObjectType($NetworkAdapter));

#Enkel bij een instantie zit hier iets in natuurlijk
printf("waarde van attribuut name: %s\n", $NetworkAdapter->{Name});

#OF

$WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");
$NetworkAdapter = $WbemServices->Get($ClassName);
printf("objecttype van klasse = %s\n", Win32::OLE->QueryObjectType($NetworkAdapter));
printf("waarde van attribuut name: %s\n", $NetworkAdapter->{Name});



