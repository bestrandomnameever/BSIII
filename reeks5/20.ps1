# Volgens mij zijn methodes die als qualifier implemented op true hebben staan

$Location = New-Object -com "WbemScripting.SWbemLocator";
$WmiService = $Location.ConnectServer(".","root\cimV2");

$oudeconfig = $ErrorActionPreference;
$ErrorActionPreference= 'silentlycontinue';

$Class = $WmiService.Get("Win32_LogicalDisk");
# Opnieuw dezelfde fout als bij 18. Als hij er niet is gaat hij fouten geven. Dat gewoon tegengaan door silentlycontinue in te stellen
$Class.Methods_ | where {$_.Qualifiers_.Item("implemented")} | select Name;

$ErrorActionPreference = $oudeconfig;