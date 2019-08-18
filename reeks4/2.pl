use Win32::OLE;

#Voor warnings
Win32::OLE->Option(Warn => 3);

$computerNaam = 'LocalHost';
$nameSpace = "root/cimv2";


$Locator = Win32::OLE->new('WbemScripting.SWbemLocator');
$WbemServices = $Locator->ConnectServer($computerNaam, $nameSpace);
print Win32::OLE->QueryObjectType($WbemServices), "\n";

#OOOOFFF, geen Locator object aanmaken:

#Dit ipv eerst een object te moeten aanmaken van de Locator: gewoon in 1 keer zo blijkt
#Werkt dit enkel voor $WbemServices??
$WbemServices = Win32::OLE->GetObject("winmgmts://$computerNaam/$nameSpace");
print Win32::OLE->QueryObjectType($WbemServices), "\n";

# Aan een ander toestel: computernaam veranderen en 4 argumenten bij de ConnectServer methode gebruiken.