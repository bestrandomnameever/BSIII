$Location = New-Object -com "WbemScripting.SWbemLocator";
$WmiService = $Location.ConnectServer(".","root\cimV2");


$AllClasses = $WmiService.ExecQuery("select * from meta_class");
$AllClasses |
    select @{Name="Naam"; Expression={$_.SystemProperties_.Item("__CLASS").Value}},
            # Ik had eerst als name Count genomen, maar dat werkt niet in de sort
           @{Name="aantal"; Expression={$_.Properties_.Count}} | sort aantal -Descending | select -first 1

