use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Win32::OLE::Variant;
use Data::Dumper;

@ARGV = ("Win32_Directory");

my $ComputerName  = ".";
my $NameSpace = "root/cimv2";

my $WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");

#WITHIN 5 is pollen om de 5 seconden
my $NotificationQuery = "SELECT * FROM __InstanceModificationEvent WITHIN 5  
                          WHERE TargetInstance ISA 'Win32_Service'";

$| = 1; # Schrijfbuffer op een waarde verschillend van 0 zodat onmiddelijk naar het scherm geschreven wordt
my $SWbemEventSource = $WbemServices->ExecNotificationQuery($NotificationQuery);

print "Waiting for event, they will be logged here: \n";

while(1){
    my $Event = $SWbemEventSource->NextEvent(5000); #Om de 5 seconden kijken
    # Als er iets inzit moeten we het printen
    $Event 
    ? printf("\n%s changed from state %s to %s\n", $Event->TargetInstance->DisplayName, $Event->PreviousInstance->State, $Event->TargetInstance->State)
    : print ".";
}
