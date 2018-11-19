$Location=New-Object -comobject "WbemScripting.SWbemLocator"

function get-Namespace ($name)
{
   $service = $Location.ConnectServer(".",$name)
   #haal alle instanties van de klasse __NAMESPACE
   $service.execQuery("select * from __NAMESPACE") | 
   foreach{
      "Namespace  = "+ $name+"\" + $_.Properties_.Item("Name").Value # Dit print gewoon de waarde. Je moet zelfs geen Write-Host gebruiken #MINDBLOWN
      get-Namespace($name+"\" + $_.Properties_.Item("Name").Value)
   }
 }
get-Namespace("root\cimV2")