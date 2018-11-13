use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Win32::OLE::Variant;
use Data::Dumper;

@ARGV = ("Win32_Directory");

my $ComputerName  = ".";
my $NameSpace = "root/cimv2";

# Refresher maken
my $Refresher=Win32::OLE->new("WbemScripting.SWbemRefresher");
my $WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");

foreach my $ClassName (@ARGV){
    # Kijken of het een singleton is. Eerst kijken of de Qualifier bestaat
    # En dan pas de value checken. Anders krijg je een fout.

    $WbemServices->Get("$ClassName")->Qualifiers_->Item("Singleton")
    && $WbemServices->Get("$ClassName")->Qualifiers_->Item("Singelton")->Value 
    ? $Refresher->Add($WbemServices, $ClassName . "=@")->{Object} # Singleton opvragen. Object is hier belangrijk om instanties op te vragen
    : $Refresher->AddEnum($WbemServices, $ClassName)->{ObjectSet}; # Dit vraagt instances op. Idem voor ObjectSet
}

# Deze oefening moeten we niet kennen dus ik ga verder met 52.


