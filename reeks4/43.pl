use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Win32::OLE::Variant;
use Data::Dumper;

my $DateTime = Win32::OLE->new("WbemScripting.SWbemDateTime");

my $ComputerName  = ".";
my $NameSpace = "root/cimv2";
my $WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");

my $ClassName =  "Win32_Service";

sub maakHashMethodeValues {
    my $Method = shift;
    my %hash = ();
    @hash{ @{$Method->Qualifiers_->{ValueMap}->Value} } = @{$Method->Qualifiers_->{Values}->Value};
    return \%hash;
}

my $class = $WbemServices->Get($ClassName, wbemFlagUseAmendedQualifiers);
%returnStartServiceHash = %{ maakHashMethodeValues($class->Methods_->{StartService}) };
%returnStopServiceHash = %{ maakHashMethodeValues($class->Methods_->{StopService}) };


# Ik vond geen Oracle Services op de VM. Besturingssytemen is een vreemd vakske
# Deze service is iets van printers, dus die kan je normaal ook afsluiten
@ARGV = ("Spooler");
my $name = shift;
my $object = $WbemServices->Get("$ClassName.Name='$name'", wbemFlagUseAmendedQualifiers);

# Als het object bestaat dan stoppen, anders starten
# StopService en StartService hebben geen parameters
# Belangrijk is de eq hier ipv ==
printf("Status of Service %s is %s\n", $name, $object->Properties_->Item("State")->Value);
if($object->Properties_->Item("State")->Value eq "Running"){
    print "Trying to stop service...\n";
    my $returnValue = $object->StopService();
    printf("Status: %s\n", $returnStopServiceHash{$returnValue});

} else {
    print "Trying to start service...\n";
    my $returnValue = $object->StartService();
    printf("Status: %s\n", $returnStartServiceHash{$returnValue});
}