$Location = New-Object -com "WbemScripting.SWbemLocator";
$WmiService = $Location.ConnectServer(".","root\cimV2");

# Hash maken om dat eens te testen
# Vrij handig in gebruik:
$hash = @{};

$Class = $WmiService.Get("Win32_LogicalDisk", 131072);
$Class.Properties_ | foreach {
    $hash[$_.Name] = $_.Qualifiers_;
}

# Printen van alle keys met zijn Qualifiers
$hash.Keys | foreach {
    Write-Host("`nProperty: " + $_);
    $hash[$_] | select Name ,Value | Format-List;
}