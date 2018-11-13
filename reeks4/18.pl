use Win32::OLE 'in';
use Data::Dumper;

# Informatie
$ComputerName = ".";
$NameSpace = "root/cimv2";
$ClassName = "Win32_Environment.Name=\"TEST\",UserName=\"<SYSTEM>\"";

# WbemServices maken
$monik = "winmgmts://$ComputerName/$NameSpace";
$WbemServices = Win32::OLE->GetObject($monik);

$environment = $WbemServices->Get($ClassName);

$environment->{VariableValue} = "Anouk";

# Opslaan van nieuwe waarde
$environment->Put_();

# Je kan nu kijken in WMI CIM STUDIO: De environment variabele value is veranderd