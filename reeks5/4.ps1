# Alle dependend services ophalen van een service
# Je kan de properties weten door Get-Service | Select -l 1 | select * te doen
Get-Service | select -l 1 | select ServicesDependedOn

# Alle services zonder dependendServices ophalen en die running zijn, enkel de naam
# Zorg zeker dat je niet eerst iets filtert en dan gaat zoeken op een property die weggefilterd is
Get-Service | where {$_.ServicesDependOn.Count -eq 0 -and $_.Status -eq "Running"} | select name