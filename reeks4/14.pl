use Win32::OLE 'in';
use Data::Dumper;

# Vraag stellen: Waarom is de SNMP service == Browser service?

# Informatie
$ComputerName = ".";
$NameSpace = "root/cimv2";
$ClassName = "Win32_Service";

# WbemServices maken
$monik = "winmgmts://$ComputerName/$NameSpace";
$WbemServices = Win32::OLE->GetObject($monik);

# Object maken 
# Zo kan je dus ook werken ipv een query
$Browser = $WbemServices->Get("$ClassName.Name='Browser'");

# Nu printen en meervoudige attributen op 1 lijn weergeven
# Belangrijk dus hier: SystemProperties is blijkaar geen Win32::OLE object
# Properties wel
for (in $Browser->Properties_, $Browser->SystemProperties_){
    print $_->{Name}, "=> ", $_->{Value}, "\n" if !$_->{IsArray};
    print $_->{Name}, " ", join ", ",  (in $_->{Value}), "\n" if $_->{IsArray};
}