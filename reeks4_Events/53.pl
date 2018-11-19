use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Win32::OLE::Variant;
use Data::Dumper;

my $ComputerName  = ".";
my $NameSpace = "root/cimv2";

my $WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");

# Met de klasse Win32_ProcessTrace kunnen we specifiek procesevents gaan bekijken
my $ClassName = "Win32_ProcessTrace";
my $Query = "select * from Win32_ProcessTrace";

# Query uitvoeren
my $EventSource = $WbemServices->ExecNotificationQuery($Query);

# Oneindige lus om te blijven checken op events
while(1){
    # NextEvent werkt blocking: zolang er geen event is, zal er niets geprint worden
    my $Event = $EventSource->NextEvent();

    # Iets heel raar. In de oplossing staat er dat we eerst gaan ophalen of het een StartTrace is of niet. 
    # Dat werkt niet bij mij dan :p
    # Op 1 lijn schrijven werkt wel gewoon
    # We gaan het ID en ProcessName (dit zijn properties uit ofwel ProcessStartTrace of ProcessStopTrace) uitschrijven
    # + Als het een start is dan printen we started, anders stopped
    printf("ProcessID: %s, Name: %s has %s\n", $Event->{ProcessID}, $Event->{ProcessName}, $Event->Path_->Class eq "Win32_ProcessStartTrace" ? "started" : "stopped");
}

