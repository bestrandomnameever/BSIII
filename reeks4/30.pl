# Door de Object Properties
# supportCreate en supportDelete

use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Data::Dumper;
$DateTime = Win32::OLE->new("WbemScripting.SWbemDateTime");

my $ComputerName  = ".";
my $NameSpace = "root/cimv2";
my $WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");

# Geeft alle klasses van de namespace
my $Objects = $WbemServices->SubclassesOf();

sub isAddableOrDeletable {
    my $Object = shift;
    $Qualifiers = $Object->{Qualifiers_};
    for (in $Qualifiers){
        return 1 if $_->{Name} =~ /create|delete/i
    }
    return 0;
}

for(in $Objects){
    my $Creatable = isAddableOrDeletable($_);
    printf("%s %s", $_->{Path_}->{Class}, $Creatable ? "kan aangemaakt / verwijderd worden" : "Kan NIET aangemaakt / verwijderd worden") if $Creatable; #Doe de if weg als je alle klassen wil zien
    
    if($Creatable){
            printf("\nMethode: ");
            if($_->{Qualifiers_}->item(CreateBy)){
            my $CreateBy = $_->{Qualifiers_}->item(CreateBy)->{Value};
            my $methode = $_->{Methods_}->Item($CreateBy);
            if(Win32::OLE::LastError()){
                print "Methode niet gevonden\n";
            }  else {
                print $CreateBy . "\n";
            }
        }
    }
}