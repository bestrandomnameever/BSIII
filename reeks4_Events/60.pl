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

# 1. Maken van EventFilter
# Iets anders doen nu, niet met extrinctevent
my $EventFilter = $WbemServices->Get("__EventFilter")->SpawnInstance_();
my $Query = qq[SELECT * FROM __InstanceCreationEvent WITHIN 1 where TargetInstance ISA 'Win32_LogicalDisk'];
$EventFilter->{Query} = $Query;
$EventFilter->{QueryLanguage} = 'WQL';
$EventFilter->Put_(wbemFlagUseAmendedQualifiers);
checkOpFouten();

# 2. Consumer
my $Consumer = $WbemServices->Get("ActiveScriptEventConsumer")->SpawnInstance_();
# Je kan een filename meegeven, maar da's niet nodig. Je kan ook rechtstreeks scripten in ScriptText
$Consumer->{ScriptText} = q[
    use Win32::OLE;
    my $shell = Win32_OLE->new("WScript.Shell");
    #TargetEvent is dus blijkbaar het event dat het veroorzaakt heeft
    my $device = $TargetEvent->{TargetInstance}->{DeviceId};  #(*)
    $shell->LogEvent(2,"USB ingeplugd op $device\n");
]
$Consumer->{ScriptingEngine} = "PerlScript";
$Consumer->Put_(wbemFlagUseAmendedQualifiers);


