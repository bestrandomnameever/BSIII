# We gaan reageren op timerevents in deze oefening
# Dit is eigenlijk oefening 39 geautomatiseerd
# Ik heb de 40/41 nog niet gedaan
use Win32::OLE qw/in EVENTS/;
use Win32::Console;
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Win32::OLE::Variant;
use Data::Dumper;

sub createWbemServices {
    my $ComputerName  = ".";
    my $NameSpace = "root/cimv2";
    return Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");
}

# Na elke put is het belangrijk om te checken op fouten
sub checkOpFouten {
    if(Win32::OLE->LastError()){
        print Win32::OLE->LastError(), "\n";
    }
}


my $WbemServices = createWbemServices();

# 1. Eventfilterklasse gaan ophalen
# Als je kijkt naar deze klasse heeft hij geen Create methode. Dus je gaat met SpawnInstance moeten werken. Feestje
# wbemFlagUseAmendedQualifiers constante kan handig zijn
my $EventFilterClass = $WbemServices->Get("__EventFilter", wbemFlagUseAmendedQualifiers);
my $EventFilterInstance = $EventFilterClass->SpawnInstance_();

# Instellen van parameters van EventFilter
# Naam moet je niet instellen, maar zo zie je snel waarover het gaat in CIM STUIDO
$EventFilterInstance->{Query} = "select * from __TimerEvent";
$EventFilterInstance->{QueryLanguage} = "WQL";

# Put gebruiken zodat de wijzigingen worden doorgevoerd
my $EventFilterPath = $EventFilterInstance->Put_();
checkOpFouten();

# 2. Consumer gaan instellen. 
# Er zijn een aantal consumers. Check de subklassen van __EventConsumer eens. Wij gebruiken een CommandLineEventConsumer
# Deze is opnieuw niet creatable dus moeten we hem met SpawnInstance_() gaan aanmaken
# Naam moet je niet instellen, maar zo zie je snel waarover het gaat in CIM STUIDO
my $CommandLineEventConsumerClass = $WbemServices->Get("CommandLineEventConsumer", wbemFlagUseAmendedQualifiers);
my $CommandLineEventConsumerInstance = $CommandLineEventConsumerClass->SpawnInstance_();

# Instellen van properties van CommandLineEventConsumer
# Dit gaat een bericht sturen naar de console

$CommandLineEventConsumerInstance->{CommandLineTemplate} = "msg console /Time: 5 'TEST'";

# Als je hier geen wbemFlagUseAmendedQualifiers meegeeft aan put gaat hij boos zijn
# De fout wordt wel opgevangen in LastError() dus je kan hier in feite niet mis zijn
my $ConsumerPath = $CommandLineEventConsumerInstance->Put_(wbemFlagUseAmendedQualifiers);
checkOpFouten();

# 3. Nu moeten we een FilterToConsumerBinding maken
# We moeten het path van de objecten gaan aanpassen zodat FilterToConsumerBinding ze begrijpt
# Als je een object aanmaakt met Put_() kan je het path niet opvragen op de normale manier
# Je moet het object dat te terugkrijgt van Put bijhouden.
# Hier moet je het path niet gaan veranderen. Blijkbaar staan die al juist na put, joepie!

my $FilterToConsumerBindingClass = $WbemServices->Get("__FilterToConsumerBinding", wbemFlagUseAmendedQualifiers);
my $FilterToConsumerBindingInstance = $FilterToConsumerBindingClass->SpawnInstance_();

# Properties setten
$FilterToConsumerBindingInstance->{Filter} = $EventFilterPath->{Path};
$FilterToConsumerBindingInstance->{Consumer} = $ConsumerPath->{Path};

$FilterToConsumerBindingInstance->Put_();
checkOpFouten();

