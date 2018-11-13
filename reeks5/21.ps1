$Location = New-Object -com "WbemScripting.SWbemLocator";
$WmiService = $Location.ConnectServer(".","root\cimV2");

# Hash maken om dat eens te testen
# Vrij handig in gebruik:
$hash = @{};

$Class = $WmiService.Get("Win32_process", 131072);

$Class.Methods_ | foreach {
    $hash[$_.Name] = $_.Qualifiers_;
}

$hash.Keys | foreach {
    Write-Host("`n`n`n`nMethod name: " + $_ );
    Write-Host("----------------------------------")
    #$hash[$_] | select Name, Value | Format-List; # Dit werkt eigenlijk perfect maar ze willen dat we Write-Host gebruiken
    $hash[$_] | foreach{ Write-Host($_.Name +  ": " + $_.Value) };
}