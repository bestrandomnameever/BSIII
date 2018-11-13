# Sleutelattributen hebben een qualifier 'key'
use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Data::Dumper;

my $ComputerName  = ".";
my $NameSpace = "root/cimv2";
my $WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");

# Geeft alle klasses van de namespace
my $Classes = $WbemServices->SubclassesOf();

sub telSleutelAttributen {
    my $Class = shift;
    my @Sleutelattributen = ();
    foreach my $Property (in $Class->Properties_){
        foreach my $Qualifier (in $Property->Qualifiers_){
            push @Sleutelattributen, $Property->{Name} if $Qualifier->{Name} =~ /key/i;
        }
    }
    # Geef een referentie terug, alejah, performance enzo
    return \@Sleutelattributen;
}

# Klasses printen met samengestelde sleutels
foreach my $Class (in $Classes){
    my @Sleutels = @{telSleutelAttributen($Class)};
    if(scalar @Sleutels > 1){
        # Naam van de klasse printen met Path_
        printf("Sleutelattributen van klasse '%s' : %s\n", $Class->{Path_}->{Class}, join ", ", @Sleutels);
    }
}

#Merk op dat een associatorklasse altijd 2 of meer sleutelattributen heeft. 
#De sleutelattributen (ref-type) bevatten de naam van de klasse die geassocieerd wordt.


