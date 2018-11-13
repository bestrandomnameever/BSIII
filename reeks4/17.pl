use Win32::OLE 'in';
use Data::Dumper;

# Informatie
$ComputerName = ".";
$NameSpace = "root/cimv2";

# WbemServices maken
$monik = "winmgmts://$ComputerName/$NameSpace";
$WbemServices = Win32::OLE->GetObject($monik);


# Seeden van klasses
# Dit is een instantie. Blijkbaar, als je dit bij een klasse doet, kloppen je keys en IsSingleton niet.
# Dat lijkt mij maar logisch in feite
@ARGV = ("Win32_Process.Handle=\"856\"");

# Properties die je wil opvragen. Deze kan je vinden in de MSDN bibliotheek
my @properties=("Authority","Class","DisplayName","IsClass", "IsSingleton","Keys","Local","Namespace","ParentNamespace", 
                "Path","RelPath","Server");

for (@ARGV){
    $class = $WbemServices->Get($_);

    $WbemObjectPath = $class->{Path_};

    for $property (@properties){
        $waarde = $WbemObjectPath->{$property};

        # Als $waarde een referentie is
        if(ref($waarde)){

            #Als waarde een set is
            if(Win32::OLE->QueryObjectType($waarde) =~ /Set/){
                print "$property => \n";
                # De hash overlopen van waarde
                for (in $waarde){
                    printf("    %s => %s\n", $_->{Name}, $_->{Value});
                }
            }
        } else {
            printf("%s => %s\n", $property, $WbemObjectPath->{$property});
        }
        
    }
}
