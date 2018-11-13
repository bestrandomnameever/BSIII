$Location = New-Object -com "WbemScripting.SWbemLocator";
$WmiService = $Location.ConnectServer(".","root\cimV2");

$Class = $WmiService.Get("Win32_Directory");



# Alle properties en system_properties ophalen
# Geen idee hoe ik die constante hash inlaadt in Powershell ==> lel, dat gaat niet
$Class.Properties_ | Select Name, IsLocal, CIMType, IsArray, Value;

$Class.SystemProperties_ | select Name, IsLocal, CIMType, IsArray, Value;