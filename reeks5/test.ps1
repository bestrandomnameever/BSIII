$Locator = New-Object -ComObject 'WbemScripting.SwbemLocator';
$Service = $Locator.ConnectServer(".", "root/cimv2");

$Class = $Service.Get("Win32_LogicalDisk", 131072);
$Class.Methods_ | foreach {
    if($_.Qualifiers_ | where {$_.Name -eq "Values"} ){
        $_.Name 
        foreach ($item in $_.Qualifiers_.Item("Values").Value){
            "   " + $item
        }
    }
}