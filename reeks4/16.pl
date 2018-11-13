use Win32::OLE 'in';
use Data::Dumper;

# Zodat argumenten worden meegegeven
@ARGV = ('Caption', 'BootDevice');

# Informatie
$ComputerName = ".";
$NameSpace = "root/cimv2";
$ClassName = "Win32_OperatingSystem";

# WbemServices maken
$monik = "winmgmts://$ComputerName/$NameSpace";
$WbemServices = Win32::OLE->GetObject($monik);

# Operating system inladen
($instance) = in $WbemServices->InstancesOf($ClassName);

for (@ARGV){
    printf("%s => %s\n", $_, $instance->Properties_($_)->{Value});
}

#Relatief pad ophalen
print $instance->SystemProperties_("__RELPATH")->{Value}, "\n";