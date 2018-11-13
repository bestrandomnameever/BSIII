$Location = New-Object -com "WbemScripting.SWbemLocator";
$WmiService = $Location.ConnectServer(".","root\cimV2");

$Instances = $WmiService.InstancesOf("Win32_LogicalDisk");
Write-Host("Aantal klassen: " +$Instances.Count);

$Instances | foreach Path_ | select RelPath 
