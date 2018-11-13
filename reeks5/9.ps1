#Analoog aan Perl eerst deze objecten maken
$Location = New-Object -com "WbemScripting.SWbemLocator";
$WmiService = $Location.ConnectServer(".","root\cimV2");

$Instances = $WmiService.InstancesOf("Win32_LogicalDisk");
foreach ($instance in $Instances){
    $Instance.Properties_ | where {$_.Name -eq "DeviceID"} | select Value
}

# of

$Instances | foreach Properties_ | where {$_.Name -eq "DeviceID"} | select Value