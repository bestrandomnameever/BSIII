$Location = New-Object -com "WbemScripting.SWbemLocator";
$WmiService = $Location.ConnectServer(".","root\cimV2");

# Om alle qualifiers te zien moet je deze shit toevoegen aan de get. Dat zal de bit zijn die wbdembadiebla
# Instelt als constante in Perl. Hier dus gewoon vanbuiten leren? Wa nen bs
$Class = $WmiService.Get("Win32_LogicalDisk", 131072);


$Qualifiers = $Class.Qualifiers_;
$Qualifiers | foreach{Write-Host($_.Name + "===> " + $_.Value)}

# Als je de value wil ophalen is deze manier soms handig
$Qualifiers.Item("Description").Value
