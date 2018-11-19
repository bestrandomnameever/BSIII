$Location = New-Object -com "WbemScripting.SWbemLocator";
$WmiService = $Location.ConnectServer(".","root\cimV2");


$Class = $WmiService.ExecQuery("Select * from Win32_Process");
$Class | sort -Property Properties_.Handle -descending | select -first 10 
# of 
Get-Process | Sort Handles -des | select -first 10