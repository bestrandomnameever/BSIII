$Location = New-Object -com "WbemScripting.SWbemLocator";
$WmiService = $Location.ConnectServer(".","root\cimV2");


# Hier een aantal mogelijkheden
$Classes = $WmiService.ExecQuery("select * from Win32_Process where Description = 'notepad.exe'");
$Classes | select @{Name='ProcesID';Expression={$_.Properties_.Item("ProcessID").Value}},@{Name='name';Expression={$_.Properties_.Item("Name").Value}}, @{Name='handlenummer';Expression={$_.Properties_.Item("Handle").Value}}