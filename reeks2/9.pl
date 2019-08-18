use Win32::OLE qw(in with);
use Data::Dumper;

my $fso = Win32::OLE->new("Scripting.FileSystemObject");
my $Excel = Win32::OLE->GetActiveObject("Excel.Application") || Win32::OLE->new("Excel.Application");
$Excel->{Visible} = 1;

sub makeArrayFromHash {
    $hash = shift;
    my $teller = 0;
    my $array;
    foreach (keys %$hash){
        $array->[$teller][0] = $_;
        $array->[$teller][1] = $hash->{$_}{punten1};
        $array->[$teller][2] = $hash->{$_}{punten2};
        $array->[$teller][3] = $hash->{$_}{gemiddelde};
        $teller++;
    }
    return $array;
}

sub makeSecondArray {
    $hash = shift;
    my $teller = 0;
    my $array;
    foreach (keys %$hash){
        $array->[$teller][0] = $_;
        $array->[$teller][1] = 'D';
        $array->[$teller][1] = 'C' if $hash->{$_}{gemiddelde} >= 7;
        $array->[$teller][1] = 'B' if $hash->{$_}{gemiddelde} >= 10;
        $array->[$teller][1] = 'A' if $hash->{$_}{gemiddelde} >= 12;
        $teller++;
    }
    return $array;
}

# Open het bestand via de pathnaam, vergeet niet te escapen
# Deze waarden gaan we opslaan in een hash
my $AbsolutePath = $fso->GetAbsolutePathName('.\Excel\punten2.xls');

my $Book = $Excel->Workbooks->Open($AbsolutePath) if $fso->FileExists('.\Excel\punten2.xls');

# Dumper flipt hier kan geen comobjecten uitschrijven hier :'(
my $Cells = $Book->Worksheets(1)->UsedRange()->Value;
#print Dumper $Cells;

# Hash vullen met punten
foreach $Cell (@$Cells){
    $hash->{$Cell->[0]}{punten2} = $Cell->[1];
}

# Ene applicatie sluiten
$Excel->Quit();

# Andere punten gaan ophalen
$AbsolutePath = $fso->GetAbsolutePathName('.\Excel\punten.xls');
$Book = $Excel->Workbooks->Open($AbsolutePath) if $fso->FileExists('.\Excel\punten.xls');

$Cells = $Book->Worksheets(1)->UsedRange()->Value;

foreach $Cell (@$Cells){
    $hash->{$Cell->[0]}{punten1} = $Cell->[1];
}

# Gemiddelde gaan berekenen
foreach $person (keys %$hash){
    $hash->{$person}{gemiddelde} = ($hash->{$person}{punten1} + $hash->{$person}{punten2}) / 2;
}

my $array = makeArrayFromHash($hash);

# Cellen gaan invullen
# Hier moet staan precies bij Value
$Book->Worksheets(1)->Range("A1:D189")->{Value} = $array;

#Voeg een sheet toe als deze nog niet bestaat
# Add gaat een sheet vooraan toevoegen in Excel!
$Book->{Worksheets}->Add() if !$Book->Worksheets("Ad Valvas");
$Book->Worksheets(1)->{Name} = "Ad Valvas" if !$Book->Worksheets("Ad Valvas");

my $other = makeSecondArray($hash);
print Dumper $other;

$Book->Worksheets(1)->Range("A1:B189")->{Value} = $other;

# Book niet vergeten saven, maar niet aangezet omdat ik wou checken
#$Book->Save();