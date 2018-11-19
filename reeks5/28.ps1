$Location = New-Object -com "WbemScripting.SWbemLocator";
$WmiService = $Location.ConnectServer(".","root\cimV2");


# Hier een aantal mogelijkheden
$Classes = $WmiService.Get("Win32_LogicalDisk.DeviceID='C:'", 131072);
$Classes | select @{Name='Catpion';Expression={$_.Properties_.Item("Caption").Value}}, @{Name='Description';Expression={$_.Properties_.Item("Description").Value}}, @{Name='FileSystem';Expression={$_.Properties_.Item("FileSystem").Value}}
# of 
#$Classes | foreach Properties_ | where {$_.Name -eq 'Caption' -or $_.Name -eq 'Description' -or $_.Name -eq 'FileSystem'} | select Value