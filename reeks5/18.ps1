$Location = New-Object -com "WbemScripting.SWbemLocator";
$WmiService = $Location.ConnectServer(".","root\cimV2");


$AllClasses = $WmiService.ExecQuery("select * from meta_class");

$oudeconfig = $ErrorActionPreference;
$ErrorActionPreference= 'silentlycontinue'
#Dit werkt maar geeft fouten omdat de qualifier "abstract" niet gevonden wordt
#Ik heb dit omzeild door errors niet te laten displayen door de globale variabele $ErrorActionPreference in te stellen.
#Of dit goed is, is nog maar de vraag
$filtered = $AllClasses | where {$_.Qualifiers_.Item("abstract")} | foreach Path_ | where {$_.Class -like "Win32*"} | select Class;
$filtered;
Write-Host("`nAantal klassen: " + $filtered.Count);

#Errors terugzetten
$ErrorActionPreference = $oudeconfig;