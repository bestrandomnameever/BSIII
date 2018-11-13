use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Data::Dumper;

my $ComputerName  = ".";
my $NameSpace = "root/cimv2";
my $WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");

$Classes = $WbemServices->SubClassesOf(undef, wbemFlagUseAmendedQualifiers);

my @Properties;

sub GetPropertiesWithOnlyValues {
    my $Class = shift;
    my $ValuesCounter = 0;
    my $ValueMapCounter = 0;
    foreach my $Property (in $Class->Properties_){
        push @Properties, $Property if $Property->Qualifiers_->Item("Values") && !$Property->Qualifiers_->Item("ValueMap");
    }
}

sub PrintValues {
    my $Property = shift;
    print $Property->{Name}, "\n";
    print '-' x 40;
    print "\n";
    @array =  @{$Property->Qualifiers_->{"Values"}->{Value}};
    for $index (0.. scalar  @array - 1){
       printf("%d => %s\n", $index , $array[$index]);
    }
    print "\n";
}

for $Class (in $Classes){
    GetPropertiesWithOnlyValues($Class);
}


# Print niet als de property er al instaat
# Ja, ik weet het , ik zou dat eerst moeten filteren maar 'k had geen zin
foreach $Property (@Properties){
    $Dubbels{$Property->{Name}}++;
    PrintValues($Property) if $Dubbels{$Property->{Name}} == 1;

}



