$Location = New-Object -com "WbemScripting.SWbemLocator";
$WmiService = $Location.ConnectServer(".","root\cimV2");

$ClassName = "Win32_Service";
$InstanceName = "Browser";

# Voor de instance
Write-Host("INSTANCE: `n");
$Instance = $WmiService.Get($ClassName + ".Name='" + $InstanceName +"'");
# Analoog aan linux' &&-operator
$($Instance.Properties_;$Instance.SystemProperties_) | select Name, Value | sort -CaseSensitive -Property Name;

Write-Host("`nCLASS: `n");

# Voor de klasse
$Class = $WmiService.Get($ClassName);
$($Class.Properties_;$Class.SystemProperties_) | select Name, Value | sort -CaseSensitive -Property Name;

