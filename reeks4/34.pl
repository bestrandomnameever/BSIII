use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting'; #inladen van de WMI constanten
use Data::Dumper;

my $ComputerName = ".";
my $NameSpace = "root/cimv2";
my $Locator=Win32::OLE->new("WbemScripting.SWbemLocator");
my $WbemServices = $Locator->ConnectServer($ComputerName, $NameSpace);

@ARGV = ("Win32_NetworkAdapter");

#my $ClassName="Win32_NetworkAdapter"; #zonder argument
@ARGV or die "Geef klassenaam vb Win32_NetworkAdapter";
my $ClassName = $ARGV[0]; 

my $Class=$WbemServices->Get($ClassName,wbemFlagUseAmendedQualifiers);
print "De attributen met ValueMap voor de klasse ",$ClassName,"\n\n";

sub getInformationForPropertyWithValuesAndValueMaps {
    my $Property = shift;

    # Als ValueMap erin zit
    # Qualifiers heeft dus rechtsreeks een valuemap die niet per se met een item moet worden opgevraagd
    if($Property->Qualifiers_->{ValueMap}){
        my @info = ($Property->Name, $Property->Qualifiers_->Item("Description")->{Value});
        my %hash;
        @hash{@{$Property->Qualifiers_->Item("ValueMap")->{Value}}} = @{$Property->Qualifiers_->Item("Values")->{Value}}; #koppelen in een hash
        push @info, \%hash;
        return \@info;
    }
    return ();
}

foreach my $Property (in $Class->{Properties_}){
    (my $Name, my $Description, my $hash) = @{getInformationForPropertyWithValuesAndValueMaps($Property)};
    if($Name){
        printf("%s\n%s\n", $Name, $Description);
        while( ($key, $value) = each(%{$hash}) ){
            printf("%s => %s\n", $key, $value);
        }
        print "\n\n";
    }
}
