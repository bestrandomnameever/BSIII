$Location = New-Object -com "WbemScripting.SWbemLocator";
$WmiService = $Location.ConnectServer(".","root\cimV2");


# Hier een aantal mogelijkheden
$Classes = $WmiService.ExecQuery("select * from Win32_Service where State = 'Stopped'");
#$Classes | foreach Properties_ | select Name
#$Classes | select @{Name='naam'; Expression={$_.Properties_.Item("Name").Value}}
# OF
$Classes | foreach {$_.Properties_.Item("Name").Value}