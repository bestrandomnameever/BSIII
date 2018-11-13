$Location = New-Object -com "WbemScripting.SWbemLocator";
$WmiService = $Location.ConnectServer(".","root\cimV2");

# Eerst alle klassen ophalen en daarna filteren lijkt mij het eenvoudigst
# Wat hier kan gebeuren
$AllClasses = $WmiService.ExecQuery("select * from meta_class");
$AllClasses | foreach Path_ | where {$_.Class -like "*Event*"} | select Class;