#oplossing die de drie reacties bevat en zeven instanties maakt.
use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting ';
$Win32::OLE::Warn = 3; #script stopt met foutmelding als er ergens iets niet lukt


my $ComputerName = ".";
my $NameSpace = "root/cimv2";
my $Locator=Win32::OLE->new("WbemScripting.SWbemLocator");
my $WbemServices = $Locator->ConnectServer($ComputerName, $NameSpace);

#periodiek
my $Instance = $WbemServices->Get("__IntervalTimerInstruction")->SpawnInstance_();
$Instance->{TimerID}  = 'test';
$Instance->{IntervalBetweenEvents} = 60000;
$instancePath=$Instance->Put_(wbemFlagUseAmendedQualifiers);

#of eenmalig
my $DateTime = Win32::OLE->new("WbemScripting.SWbemDateTime"); #data manipuleren
$Instance = $WbemServices->Get("__AbsoluteTimerInstruction")->SpawnInstance_();
$Instance->{TimerID}  = 'test2';
$DateTime->SetVardate("09/11/2018 11:00:00"); #het gewenste moment
$Instance->{EventDateTime} = $DateTime->{Value};
$instancePath=$Instance->Put_(wbemFlagUseAmendedQualifiers);

#Oorzaak event koppelen aan voorgaand object : 
my $InstanceEvent = $WbemServices->Get("__EventFilter")->SpawnInstance_();
$InstanceEvent->{Name}="test";
$InstanceEvent->{QueryLanguage}="WQL";
$InstanceEvent->{Query}="SELECT * FROM __TimerEvent where TimerID = 'test2'"; #wijzig dit naar test indien je periodieke reacties wilt
$Filter = $InstanceEvent->Put_(wbemFlagUseAmendedQualifiers);
$Filterpad=$Filter->{path};

#gewenste melding
$melding = "timer wordt gelogd op tijdstip %TIME_CREATED% op toestel %__SERVER%";

#Reactie1
my $InstanceReactie1 = $WbemServices->Get("CommandLineEventConsumer")->SpawnInstance_();
$InstanceReactie1->{Name}="test";
$InstanceReactie1->{CommandLineTemplate} =  "msg console /Time:5 $melding";
$Consumer = $InstanceReactie1->Put_(wbemFlagUseAmendedQualifiers);
$Consumerpad1=$Consumer->{path};  

#Reactie2
my $InstanceReactie2= $WbemServices->Get("LogFileEventConsumer")->SpawnInstance_();
$InstanceReactie2->{Name}="test";
$InstanceReactie2->{FileName} = 'C:\\\\temp\\log.txt';
$InstanceReactie2->{Text} = "timer wordt gelogd op tijdstip $melding";
$Consumer = $InstanceReactie2->Put_(wbemFlagUseAmendedQualifiers);
$Consumerpad2=$Consumer->{path};  

#Reactie3
my $InstanceReactie3= $WbemServices->Get("NTEventLogEventConsumer")->SpawnInstance_();
$InstanceReactie3->{Name}="test";
$InstanceReactie3->{SourceName} = "WinMgmt";
$InstanceReactie3->{EventType} = 0;
$InstanceReactie3->{EventID} = 2017;#0xC0000A;
$InstanceReactie3->{NumberOfInsertionStrings} = 1;
$InstanceReactie3->{Category} = 0;
$InstanceReactie3->{InsertionStringTemplates }=[$melding];  #als array doorgeven
$Consumer = $InstanceReactie3->Put_(wbemFlagUseAmendedQualifiers);
$Consumerpad3=$Consumer->{path};  

#3 koppelingen:
my $InstanceKoppeling1 = $WbemServices->Get("__FilterToConsumerBinding")->SpawnInstance_();
$InstanceKoppeling1->{Filter}   = $Filterpad; 
$InstanceKoppeling1->{Consumer} = $Consumerpad1;
$Result=$InstanceKoppeling1->Put_(wbemFlagUseAmendedQualifiers);
print "\n1.\n",$Result->{Path},"\n"; #ter controle

my $InstanceKoppeling2 = $WbemServices->Get("__FilterToConsumerBinding")->SpawnInstance_();
$InstanceKoppeling2->{Filter}   = $Filterpad; 
$InstanceKoppeling2->{Consumer} = $Consumerpad2;
$Result=$InstanceKoppeling2->Put_(wbemFlagUseAmendedQualifiers);
print "\n2.\n",$Result->{Path},"\n"; #ter controle

my $InstanceKoppeling3 = $WbemServices->Get("__FilterToConsumerBinding")->SpawnInstance_();
$InstanceKoppeling3->{Filter}   = $Filterpad; 
$InstanceKoppeling3->{Consumer} = $Consumerpad3;
$Result=$InstanceKoppeling3->Put_(wbemFlagUseAmendedQualifiers);
print "\n3.\n",$Result->{Path},"\n"; #ter controle

#opmerking - je kan het tijdstip berekenen vanaf het huidig moment:
my ($sec,$min,$hour,$dag,$maand,$jaar) = localtime(time+5); #over 5 seconden
  $jaar+=1900;
  $maand++;
$DateTime->SetVardate("$dag/$maand/$jaar $hour:$min:$sec"); #het gewenste moment