use Win32::OLE;

#Constante ClassName
$ClassName = "__NAMESPACE";

sub GetNameSpaces {
    #BELANGRIJK:  my niet vergeten hier,zodat de variabele niet out of scope gaat
    my ( $NameSpace ) = @_;
    # Informatie
    my $ComputerName = ".";

    # monik
    $monik = "winmgmts://$ComputerName/$NameSpace";

    #WbemServices laden
    my $WbemServices = Win32::OLE->GetObject($monik);

    #Krijg de __NAMESPACE klasses
    my $ClassNameObject = $WbemServices->get($ClassName);

    # Tel aantal backslashes om zo het level te bekomen
    my @aantalBackslashes = $NameSpace =~ /\//g;

    #Dit zorgt ervoor dat het inspringt. Die x gaat een aantal keer een teken printen
     print "|";
     print "-" x scalar @aantalBackslashes*5;

     # NIEUW: aantallen bepalen:
     my $aantal = $WbemServices->ExecQuery("select * from meta_class")->Count;
     my $onmiddellijkAantal = $WbemServices->SubclassesOf(undef, 1)->Count;
     printf("> %s aantal: %d, onmiddellijke: %d\n", $NameSpace, $aantal, $onmiddellijkAantal);

    my %instanceNames = ();
    my $instances = $ClassNameObject->Instances_();
    $instanceNames{$_->{Name}} = 1 for in $instances;

    for my $name (sort { "\L$a" cmp "\L$b" }  keys %instanceNames){
        $samengesteldenaam="$NameSpace\/" . $name;
        GetNameSpaces($samengesteldenaam);
    }
}

# Als je enkel van StdRegProv moet je ze nog filteren op RELPATH ofzoiets.

GetNameSpaces("root");