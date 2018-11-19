# Met de standaard event klasse
use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Win32::OLE::Variant;
use Data::Dumper;

my $ComputerName  = ".";
my $NameSpace = "root/cimv2";

my $WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");

# Dit is een subklasse van Event. Je kan dit gewoon gaan bekijken in CIM studio
# Subklasses van deze klasse zijn InstanceOperationCreateEvent en InstanceOperationDeleteEvent
# Deze moeten we dus gaan bekijken
my $ClassName = "__InstaceOperationEvent";

# Hier moet within (om de hoeveel tijd hij events gaat ophalen) staan, anders werkt de query niet. Test deze eens uit in wbemtest
# We willen enkel processen vinden
# TargetInstance is een property van __InstanceOperationEvent
# Subklassen van __InstanceOperationEvent zijn modification, invocation, creation en deletion
# Ik dacht om te filteren op __CLASS = create en delete. Maar systemproperties kan je blijkbaar niet gaan filteren :'(. Dat moet dus achteraf
my $Query = "select * from $ClassName within 1 where TargetInstance ISA 'Win32_Process'";

# Query uitvoeren
my $EventSource = $WbemServices->ExecNotificationQuery($Query);

while(1){
    my $Event = $EventSource->NextEvent();
    my $className=$Event->{Path_}->{Class};
    next if  $className eq "__InstanceModificationEvent";  # Skippen als modification
    next if  $className eq "__InstanceInvocationEvent"; # Skippen als invocation
  
    printf "Process %-29s (%s)started \n", $Event->{TargetInstance}->{Name}, 
                  $Event->{TargetInstance}->{Handle} if  $className eq "__InstanceCreationEvent";
    printf "Process %-29s (%s)stopped \n", $Event->{TargetInstance}->{Name},
                  $Event->{TargetInstance}->{Handle} if  $className eq "__InstanceDeletionEvent"; 
}