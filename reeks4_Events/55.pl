use Win32::OLE qw/in EVENTS/;
use Win32::Console;
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Win32::OLE::Variant;
use Data::Dumper;

sub EventCallBack {
    #argumenten zijn: bronobject die het event afvuurt, naam van het event, eventobject
    my ($Source,$EventName,$Event) = @_;


    return unless $EventName eq "OnObjectReady";
    # $EventName kan volgens de documentatie gelijk zijn aan
    # Wel niet gevonden bij $EventName, wel bij het SWBEMSink Object. Die heeft methodes die deze properties bevatten
    
    # OnCompleted: Triggered when an asynchronous operation is complete.
    # OnObjectPut: Triggered when an asynchronous Put operation is complete.
    # ObObjectReady: Triggered when an object provided by an asynchronous call is available.
    # OnProgress: Triggered to provide the status of a asynchronous operation.
    
    # ClassName opvragen
    my $className=$Event->{Path_}->{Class};

    # modification moeten we niet hebben
    return if  $className eq "__InstanceModificationEvent";

    # Invocation ook niet
    return if  $className eq "__InstanceInvocationEvent";
    
    # Voor de eerste query
    if ($Event->{TargetInstance}->{Path_}->{Class} eq "Win32_Process") {
       printf "%-29s started\n", $Event->{TargetInstance}->{Name} if  $className eq "__InstanceCreationEvent";
       printf "%-29s stopped\n", $Event->{TargetInstance}->{Name} if  $className eq "__InstanceDeletionEvent";
    }

    # Voor de tweede
    else	{
       printf "%-29s %s\n", $Event->{TargetInstance}->{Path_}->{Class}, $Event->{TargetInstance}->{Path};
    }
}

# Progid gevonden door te zoeken in regedit
my $Sink = Win32::OLE->new ('WbemScripting.SWbemSink');
Win32::OLE->WithEvents($Sink,\&EventCallBack);  #koppel de gewenste methode aan dit object, let op de \&

# Initialiseren van $WbemServices
my $ComputerName  = ".";
my $NameSpace = "root/cimv2";

my $WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");


my $Query1 = "SELECT * FROM __InstanceOperationEvent WITHIN 5 WHERE TargetInstance ISA 'Win32_Process'";
$WbemServices->ExecNotificationQueryAsync($Sink, $Query1); 


# Initialiseren van tweede $WbemService want we moeten een andere namespace hebben
$ComputerName  = ".";
$NameSpace = "root/msapps12";

my $WbemServices2 = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");

my $Query2 = "SELECT * FROM __InstanceCreationEvent WITHIN 5 WHERE TargetInstance ISA 'Win32_ExcelWorkBook'";
$WbemServices2->ExecNotificationQueryAsync($Sink, $Query2);

my $Console = new Win32::Console(STD_INPUT_HANDLE);  #creeert een nieuw Console object
$Console->Mode( ENABLE_PROCESSED_INPUT);    #enkel reageren op toetsen, niet op muis-bewegingen

 $|=1; 
 print "Wacht op een actie\n.";

until ($Console->Input()) {   #zolang er geen input is
      # Deze methode is noodzakelijk, anders gaat de callback niet uitgevoerd worden
      Win32::OLE->SpinMessageLoop();
      Win32::Sleep(500);
}


