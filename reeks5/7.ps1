#Analoog aan Perl eerst deze objecten maken
$Location = New-Object -com "WbemScripting.SWbemLocator";
$WmiService = $Location.ConnectServer(".","root\cimV2");

$Class = $WmiService.Get("Win32_LogicalDisk");

# Alle properties tonen
$Class.Properties_;
$Class.SystemProperties_;

# Waarde tonen van __DERIVATION
# klassieke mannier
$Class.SystemProperties_.Item("__DERIVATION").Value;

# Manier zoals je in Powershell kan doen
$Class.SystemProperties_ | where {$_.Name -eq "__DERIVATION"} | select Value;

# Met enkel powershell functies. Kan je zo doen maar niet de bedoeling om op de test te gebruiken
$Class = Get-WmiObject -Class "Win32_LogicalDisk";
$Class | select *;

