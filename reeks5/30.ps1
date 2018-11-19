$Location = New-Object -com "WbemScripting.SWbemLocator";
$WmiService = $Location.ConnectServer(".","root\cimV2");


$AllClasses = $WmiService.ExecQuery("select * from meta_class");
$AllClasses | select @{Name='ClassName'; Expression={$_.Path_.Class}} | where {$_ -like '*Win32**NetWorkAdapter*'}
# Lees dit om het verschil te snappen tussen -contains en -like
# https://www.itprotoday.com/powershell/powershell-contains