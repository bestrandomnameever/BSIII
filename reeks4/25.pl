use Win32::OLE 'in';
use Data::Dumper;

#Constante ClassName
$ClassName = "__NAMESPACE";
$NameSpace = "root\\CIMV2";
$ComputerName = ".";
# monik
$monik = "winmgmts://$ComputerName/$NameSpace";

#WbemServices laden
my $WbemServices = Win32::OLE->GetObject($monik);

#Krijg de __NAMESPACE klasses
my $ClassNameObject = $WbemServices->get($ClassName);

sub GetAllClasses {
     my $klassen = $WbemServices->SubclassesOf(undef, 1);
     printf("> %s met volgende klassen\n", $NameSpace);
     printKlassesRecursief($klassen, 0);
}

sub printKlassesRecursief{
    (my $klassen, my $niveau) = @_;
         for (in $klassen){
            my $WbemObjectPath = $_->{Path_};
            my $ClassName =  $WbemObjectPath->{Class};
            
            print "-" x ($niveau * 10);
            printf(">%s (niveau $niveau)\n", $ClassName);
            my $nieuweKlassen = $WbemServices->SubclassesOf($ClassName, 1);

            #Niveau gaat recursief groter worden
            #Je zou hier ook met de derivation property kunnen werken
            printKlassesRecursief($nieuweKlassen, $niveau + 1);
     }
}

# Als je enkel van StdRegProv moet je ze nog filteren op RELPATH ofzoiets.

GetAllClasses();