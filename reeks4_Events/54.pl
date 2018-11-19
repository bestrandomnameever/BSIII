# 54.1 al gedaan in 53? 
# We doen het met Office programma's

use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Win32::OLE::Variant;
use Data::Dumper;

my $ComputerName  = ".";

# Hierdoor werkte het niet. Blijkbaar zitten de Microsoft Apps in deze namespace. Amai dat duurde lang om te vinden
my $NameSpace = "root/msapps12";

my $WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");

# Met de klasse Win32_ProcessTrace kunnen we specifiek procesevents gaan bekijken
my $ClassName = "__InstanceOperationEvent";

# Query uitvoeren
my $EventSource = $WbemServices->ExecNotificationQuery("SELECT * FROM __InstanceCreationEvent WITHIN 5 WHERE TargetInstance ISA 'Win32_ExcelWorkBook'");
print Win32::OLE->LastError();

while(1){
	my $Event = $EventSource->NextEvent();
	printf "%-29s %s\n", $Event->{TargetInstance}->{Path_}->{Class}, $Event->{TargetInstance}->{Path};
}