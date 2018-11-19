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

my $WbemServices = createWbemServices();

# 1. Eventfilter
my $EventFilter = $WbemServices->Get("__EventFilter")->SpawnInstance_();
$EventFilter->{Name} = 'test';
# Hier moet je nooit () gebruiken. Enkel bij meervoudige and en or statements
$EventFilter->{Query} = qq[select * from __InstanceCreationEvent within 5 where TargetInstance ISA 'Win32_Process' AND TargetInstance.Name = 'notepad.exe'];
$EventFilter->{QueryLanguage} = "WQL";

my $EventFilterPath = $EventFilter->Put_(wbemFlagUseAmendedQualifiers);
Win32::OLE->LastError();

# 2. Consumer
my $Consumer = $WbemServices->Get("ActiveScriptEventConsumer")->SpawnInstance_();
$Consumer->{Name} = "test";
$Consumer->{ScriptingEngine} = "PerlScript";

# TODO: Wat doet die Debug weer in dat script bij het aanmaken van $WbemServices?
# Dit script gaat gewoon kijken naar de creationdate van de targetinstance
# Alle creationdates die hij kan ophalen na degene waardoor het event getriggered is
# moet hij afsluiten
# WIn32_Process heeft wel een methode Terminate
$Consumer->{ScriptText} = q[use Win32::OLE 'in';
  my $WbemServices = Win32::OLE->GetObject("winmgmts:{(Debug)}!//./root/cimv2");
  $_->Terminate foreach in $WbemServices->ExecQuery("Select * From Win32_Process
	 Where Name='$TargetEvent->{TargetInstance}->{Name}'
	and Creationdate<'$TargetEvent->{TargetInstance}->{Creationdate}'"); ];
my $ConsumerPath = $Consumer->Put_(wbemFlagUseAmendedQualifiers);
Win32::OLE->LastError();

# 3. EvenToConsumerBinder
my $FilterToConsumer = $WbemServices->Get("__FilterToConsumerBinding")->SpawnInstance_();
$FilterToConsumer->{Filter} = $EventFilterPath->{Path};
$FilterToConsumer->{Consumer} = $ConsumerPath->{Path};
Win32::OLE->LastError();

# Melding toevoegen wanneer gestart
my $Consumer2 = $WbemServices->Get("ActiveScriptEventConsumer")->SpawnInstance_();
$Consumer2->{Name} = "message";
$Consumer2->{ScriptingEngine} = "PerlScript";
$Consumer2->{ScriptText} = qq[
    use Win32::OLE;
    my $Shell = Win32:OLE->new("WScript.Shell");
    my $Name = $TargetEvent->{TargetInstance}->Name;
    # Je weet dat de TargetInstance van het type Win32_Process moet zijn. Deze heeft een creation date
    my $date = $TargetEvent->{TargetInstance}->Creationdate;
    $Shell->LogEven(2, "$Name gestart op $date\n");
]
my $Consumer2Path = $Consumer2->Put_(wbemFlagUseAmendedQualifiers);

# Tweede koppeling
my $FilterToConsumer2 = $WbemServices->Get("__FilterToConsumerBinding")->SpawnInstance_();
$FilterToConsumer2->{Filter} = $EventFilterPath->{Path};
$FilterToConsumer2->{Consumer} = $Consumer2Path->{Path};
Win32::OLE->LastError();

# Als er gegroepeert wordt in de query met GROUP WITHIN dan is er een representative van de groep objecten ==> het eerste object.
# In de oplossing tonen ze het met die representative ook


