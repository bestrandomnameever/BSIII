use Win32::OLE 'in';
use Data::Dumper;
$ComputerName = ".";
$NameSpace = "root/cimv2";
# Zoeken op de propery SystemVariable zodat je op deze klasse uitkomt
$ClassName = "Win32_Environment";

#WbemService maken
$monik = "winmgmts://$ComputerName/$NameSpace";
$Wbemservices = Win32::OLE->GetObject($monik);

#Voor systeemvarbiabelen
#Klasse maken
$InstancesSet = $Wbemservices->ExecQuery("select * from $ClassName where SystemVariable=TRUE");

#Overlopen
for (sort {lc($a->{Name}) cmp lc($b->{Name})} (in $InstancesSet)){
    printf("NAME: %20s value: %20s, user: %20s\n", $_->{Name}, $_->{VariableValue}, $_->{UserName});
}