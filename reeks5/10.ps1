$Location = New-Object -com "WbemScripting.SWbemLocator";
$WmiService = $Location.ConnectServer(".","root\cimV2");

$class = $WmiService.Get("Win32_LogicalDisk");
# Wreed raar maar je moet hier zorgen dat je foreach gebruikt ipv select. Select Qualifiers_ geeft u het object zelf terug
$Properties = $class.Properties_;
foreach ($Property in $Properties){
    if($Property | foreach Qualifiers_ | where {$_.IsLocal.ToString() -eq "True"}){
        $Property | select Name;
    } 
    
} 