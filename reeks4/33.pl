# Associatorklasse bevat een recursieve relatie als de klassenaam van het Antecedent en Dependent hetzelfde zijn
# Niet direct een idee hoe anders

use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Data::Dumper;

my $ComputerName  = ".";
my $NameSpace = "root/cimv2";
my $WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");

# Methode om te kijken of iets geset is
sub isSetTrue{
  my ($Object,$prop)=@_;
  return  $Object->{Qualifiers_}->Item($prop) && $Object->{Qualifiers_}->Item($prop)->{Value};
}

sub classHasProperty {
    my ($Object,$prop)=@_;
    return $Class->Properties_->item($prop);
}

# Geeft alle klasses van de namespace
my $Classes = $WbemServices->SubclassesOf();

sub isRecursiveAssociator {
    my $Class = shift;

    # Enkel verdergaan wanneer de klasse een associatorklasse is.
    next unless isSetTrue($Class, "Association");

    # Geen abstracte klassen gaan controleren
    next if isSetTrue($Class, "Abstract");

    # Enkel 1 returnen als Antecedentklasse gelijk is aan Dependentklasse.
    # Controleer eerst of de klasse een item bevat dat Antecedent bevat want niet elke klasse heeft dat
    return 1 if classHasProperty($Class, "Antecedent") && classHasProperty($Class, "Dependent") && $Class->Properties_->Item('Antecedent')->{Type} == $Class->Properties_->Item('Dependent')->{Type};
}

for $Class (in $Classes){
    printf("%s is een recursieve associatorklasse\n", $Class->{Path_}->{Class}) if isRecursiveAssociator($Class);
}



