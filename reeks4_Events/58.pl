# Een usb proberen inpluggen
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

# 1. Eventfilterklasse
my $EventFilterClass = $WbemServices->Get("__EventFilter", wbemFlagUseAmendedQualifiers);
my $EventFilterInstance = $EventFilterClass->SpawnInstance_();

# Van welk soort Event is het insteken van een USB?
# Ik zoek in CIM Studio gewoon naar 'Event'. 
# Alternatief is misschien om de klasse __Event en deze open te klappen
# Ik vond Win32_DeviceChangeEvent. Die geef ik dan in in wbemtest
# select * from Win32_DeviceChangeEvent
# Dan trek ik de USB stick uit of steek ik hem in en zie een event van type Win32_VolumeChangeEvent

$EventFilterInstance->{Name} = "58.PLFILTER";
$EventFilterInstance->{Query} = qq[SELECT * FROM Win32_VolumeChangeEvent];
$EventFilterInstance->{QueryLanguage} = "WQL";

my $EventFilterPath = $EventFilterInstance->Put_();
checkOpFouten();

# 2. Consumer, eerst een CommandLineConsumer
my $CommandLineEventConsumerClass = $WbemServices->Get("CommandLineEventConsumer", wbemFlagUseAmendedQualifiers);
my $CommandLineEventConsumerInstance = $CommandLineEventConsumerClass->SpawnInstance_();

# Blijkbaar maakt ge hier een message aan
my $message =  q[USB (%DriveName%) inserted op %__SERVER% - %__CLASS%];; 

$CommandLineEventConsumerInstance->{CommandLineTemplate} = "msg console /Time: 5 $message";
$CommandLineEventConsumerInstance->{Name} = "58.PLCONSUMER";

my $ConsumerPath = $CommandLineEventConsumerInstance->Put_(wbemFlagUseAmendedQualifiers);
checkOpFouten();

# 3. Binding
my $FilterToConsumerBindingClass = $WbemServices->Get("__FilterToConsumerBinding", wbemFlagUseAmendedQualifiers);
my $FilterToConsumerBindingInstance = $FilterToConsumerBindingClass->SpawnInstance_();

# Properties setten
$FilterToConsumerBindingInstance->{Filter} = $EventFilterPath->{Path};
$FilterToConsumerBindingInstance->{Consumer} = $ConsumerPath->{Path};

$FilterToConsumerBindingInstance->Put_();
checkOpFouten();