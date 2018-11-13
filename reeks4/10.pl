use Win32::OLE 'in';

# Ik begrijp niet goed wat ze hier bedoelen met het verschil tussen deze twee

#Informatie
$ComputerName = ".";
$NameSpace = "root/cimv2";

# Dit kan ook anders
$ClassName = "Win32_LogicalDisk.DeviceID=\"C:\"";

# monik
$monik = "winmgmts://$ComputerName/$NameSpace";

# WbemServices laden
$WbemServices = Win32::OLE->GetObject($monik);

# Krijg de klasse
$LogicalDiskC = $WbemServices->get($ClassName);

# Met Associators klasse proberen
# Enkel instanties gaan opvragen met 5" bit op true te zetten. Dit is de bClassesOnly bit
$Associators = $WbemServices->AssociatorsOf($ClassName);

printf("Er zijn %d instantie-associaties van $ClassName\n", $Associators->Count);

# Enkel klasses gaan ophalen
$SchemaAssociators = $LogicalDiskC->Associators_(undef, undef, undef, undef, 1);

printf("Er zijn %d verschillende klasseschema's bij de associaties van $ClassName\n", $SchemaAssociators->Count);
