use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Data::Dumper;

my $ComputerName = ".";
my $WbemServices =  Win32::OLE->GetObject("winmgmts://$ComputerName/root/cimv2");

# Alle subclasses opvragen van de root
# Door de wbemFlagUseAmendedQualifiers worden alle qualifiers opgehaald. Anders gebeurt dat niet.
my $Classes = $WbemServices->SubclassesOf(undef,wbemFlagUseAmendedQualifiers);
# Kan ook met een query, meta_class gaat de qualifiers beschrijven
#my $Classes = $WbemServices->Execquery("select * from meta_class","WQL",wbemFlagUseAmendedQualifiers);

# Overloop alle klasses. Als de provider niet is opgegeven tellen we een value bij in een hash
for $Class (in $Classes){
    if($Class->{Qualifiers_}->Item("Provider")){
        # Eentje bijtellen
        $providers->{$Class->{Qualifiers_}->Item("Provider")->{Value}}++;
    } else {
        # Ontbreekt erbij zetten
        $providers->{"Ontbreekt"}++;
    }
}

foreach ( sort { $providers->{$b} <=> $providers->{$a} } keys %{$providers} ){
    printf("%40s met value %d\n", $_, $providers->{$_});
}