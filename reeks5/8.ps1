#Analoog aan Perl eerst deze objecten maken
$Location = New-Object -com "WbemScripting.SWbemLocator";
$WmiService = $Location.ConnectServer(".","root\cimV2");

$Instance = $WmiService.InstancesOf("Win32_LogicalDisk") | select -f 1;
$Instance.Properties_.Item("DeviceID").Value;
# of 
$Instance.Properties_ | where {$_.Name -eq "DeviceID" -or $_.Name -eq "VolumeName" -or $_.Name -eq "Description"};