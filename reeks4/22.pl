use Win32::OLE 'in';
use Win32::OLE::Variant;
my $DateTime = Win32::OLE->new("WbemScripting.SWbemDateTime");

use Data::Dumper;

my $ComputerName  = ".";
my $NameSpace = "root/cimv2";
my $WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");
my $ClassName = "Win32_Service";
my $Query = "Select * From $ClassName Where State ='Running'";

# Instanties die runnen ophalen
my $Instances = $WbemServices->ExecQuery($Query);

sub print_relevante {
    (my $object) = shift;
    @relevante = ('Name', 'Description');
    foreach(@relevante){
        printf("%s = %s\n", $_, $object->{$_});
    }
}


# Dit kan performanter, ze halen alle dependentservices op en gaan dan gaan filteren op antecedent en decendent
sub print_dependencies {
    (my $object) = shift;
    my $Antecedent = "Win32_Service.Name=\"" . $object->{Name} . "\"";
    # Deze query beter studeren voor de test
    my $DependencyQuery="Associators of {$Antecedent} Where AssocClass=Win32_DependentService Role=Dependent";
    my $dependencies = $WbemServices->ExecQuery($DependencyQuery);
    print "DEPENDENCIES:\n";
    foreach(in $dependencies){
        printf("%s\n", $_->{Name});
    }
}


# Nu relevante properties gaan opschrijven en dependencies gaan ophalen
for (in $Instances){
    print_relevante($_);
    print_dependencies($_);
    print "\n\n";
}
