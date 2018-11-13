# Je kan met wbemtest methodes van een klasse gaan bekijken met zijn inputparameters en outputparameters.

use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Win32::OLE::Variant;
use Data::Dumper;

my $DateTime = Win32::OLE->new("WbemScripting.SWbemDateTime");

my $ComputerName  = ".";
my $NameSpace = "root/cimv2";
my $WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");

my $ClassName =  "Win32_Process";
my $Class = $WbemServices->Get($ClassName, wbemFlagUseAmendedQualifiers);

# Nodig om een valuemap te maken om de returnwaarde te kennen
my $Methods = $WbemServices->Get($ClassName,wbemFlagUseAmendedQualifiers)->{Methods_};

sub maakHash_method_qualifier{
   my $Method=shift;
   my %hash=();
   @hash{@{$Method->Qualifiers_(ValueMap)->{Value}}}
        =@{$Method->Qualifiers_(Values)->{Value}};
   return %hash;
}

# Caption gevonden in WMI CIM STUDIO
my $Query = "select * from $ClassName where Name='notepad.exe'";
my $Instances = $WbemServices->ExecQuery($Query);

my %TerminateReturnValues=maakHash_method_qualifier($Methods->Item("Terminate"));

print Dumper $Instances;

# Directe methode
#for my $Instance (in $Instances){
#    my $ReturnValue = $Instance->Terminate();
#    print $ReturnValue,": ",$TerminateReturnValues{$ReturnValue},"\n";
#}



# Indirecte methode
# De input parameters moeten ingevuld worden omdat er een parameter Reason is dat verplicht
# Ingevuld dient te worden. Je mag ze wel leeg laten
for my $Instance (in $Instances){
    $InputParameters = $Methods->{"Terminate"}->{InParameters};
    $Instance->ExecMethod_(Terminate, $InputParameters);
}