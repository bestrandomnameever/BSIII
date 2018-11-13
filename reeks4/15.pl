use Win32::OLE 'in';
use Data::Dumper;

# Date ophalen om datum te kunnen parsen
my $DateTime = Win32::OLE->new("WbemScripting.SWbemDateTime");

# Informatie
$ComputerName = ".";
$NameSpace = "root/cimv2";
$ClassName = "Win32_OperatingSystem";

# WbemServices maken
$monik = "winmgmts://$ComputerName/$NameSpace";
$WbemServices = Win32::OLE->GetObject($monik);

# Operating system inladen
($instance) = in $WbemServices->InstancesOf($ClassName);

for (in $instance->{Properties_}, $instance->{SystemProperties_}){
    
    #Vraag, hoe weten we dat 101 van het type datum is? Door te gaan kijken in MSDN
    #kijken bij WbemCimTypeEnum
    if ($_->{CIMType} == 101){ #datum, GetVarDate werkt niet bij mij. GetVarDate geeft mij precies een object terug
       $DateTime->{Value} = $_->{Value};     
       #print $_->{Name}, "=>" ,$DateTime->{GetVarDate}, "\n";
       # Zo werkt het wel
       print $_->Name, "=> ",$DateTime->Day, "/", $DateTime->Month, "/", $DateTime->Year, "\n";
    } else {
        print $_->{Name}, "=> ", $_->{Value}, "\n" if !$_->{IsArray};
        print $_->{Name}, " ", join ", ",  (in $_->{Value}), "\n" if $_->{IsArray};
    }
}