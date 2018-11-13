get-command -name *service* -CommandType "Cmdlet"; #Krijg alle commando's met naam service in
# Het zal wellicht om de cmdlet Get-Service gaan

$services = Get-Service;
$services.Count;

# of 
(Get-Service).Count;

# Bepalen van alle properties van één enkele service
Get-Service | select -First 1 | select *;

# Bepalen van alle services die stopped zijn
get-service |  Where {$_.status -eq "Stopped"}


