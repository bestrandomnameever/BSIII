use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Data::Dumper;

my $ComputerName  = ".";
my $NameSpace = "root/cimv2";
my $WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");
my $ClassName ="Win32_Directory";

my $Object = $WbemServices->Get($ClassName);

# Inladen van constanten
my $typeLib = Win32::OLE::Const->Load($WbemServices);

# Manier om de hash op te stellen van de CimTypes
# Vanbuiten leren dus
while( ($key, $value) = each %{$typeLib} ){
    $cimtypes{$value} = substr($key,11) if $key =~ /wbemCim/;
}

# Om eens te kijken wat er juist in de hash zit.
print Dumper \%cimtypes;

# Methode die zowel een array kan teruggeven als een gewone value
sub giveValue {
    my $Qualifier = shift;
    return $Qualifier->{Value} =~ /ARRAY/i ? join ", ", @{$Qualifier->{Value}} : $Qualifier->{Value};
}

#Elke property overlopen, daarin elke qualifier overlopen
foreach (in $Object->Properties_){
    printf("'%s' met volgende attributen: \n", $_->{Name});
    print '-' x 40;
    print "\n";
    foreach my $Qualifier (in $_->Qualifiers_){
        printf("%s => %s\n", $Qualifier->{Name}, giveValue($Qualifier));
        printf("CIMTYPE VOLGENS PROPERTY: %s\n", $cimtypes{$_->{CIMType}}) if $Qualifier->{Name} =~ /cim/i; 
    }
    print "\n\n";
}