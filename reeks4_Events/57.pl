use Win32::OLE qw/in EVENTS/;
use Win32::Console;
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Win32::OLE::Variant;
use Data::Dumper;

# 1. Verwijderen van EventFilters
sub createWbemServices {
    my $ComputerName  = ".";
    my $NameSpace = "root/cimv2";
    return Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");
}

sub deleteInstances {
    my $Instances = shift;
    foreach (in $Instances){
        # Aangezien alle $instances hier geen delete methode hebben moeten we Delete_ gebruiken.
        # Zie oefening 50
        $_->Delete_;

        # Check op fouten
        Win32::OLE->LastError();
    }
}

my $WbemServices = createWbemServices();

my $EventFilterInstances = $WbemServices->InstancesOf("__Eventfilter", wbemFlagUseAmendedQualifiers);
my $ConsumerInstances = $WbemServices->InstancesOf("CommandLineEventConsumer", wbemFlagUseAmendedQualifiers);
my $FilterToConsumerBindingInstances = $WbemServices->InstancesOf("__FilterToConsumerBinding", wbemFlagUseAmendedQualifiers);

my @array = ($EventFilterInstances, $ConsumerInstances, $FilterToConsumerBindingInstances);

foreach (@array){
    deleteInstances($_);
}