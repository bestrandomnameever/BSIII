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

# Aantal programma's die moeten openen
my $Aantal = 3;

# In WMI CIM Studio kan je vinden wat de parameters van een methode zijn 
# De pijlen duiden aan of het gaat om input of output parameters
# Je kan een methode testen (toch een statische) door rechtermuisknop execute method te doen
# Da's bet handig als je niet weet wat de parameters zijn want 't is precies wat gokken soms
# De methode zou procesid's moeten teruggeven
for (0..$Aantal-1){
    my $InParams = $Class->Methods_->{Create}->InParameters;
    $InParams->{CommandLine} = "notepad.exe";
    $InParams->{CurrentDirectory} = ".";
    $InParams->{ProcessStartupInformation} = undef;
    
    # Controle parameters
    foreach(in $InParams->Properties_){
        print Dumper $_;
    }
    $OutParameters = $Class->ExecMethod_("Create", $InParams); 
    push @Handles, $OutParameters->{ProcessId};
}

print Dumper \@Handles;





# Sleep for 3 seconds
sleep 1;

# Terug alles afsluiten
# Het lukt bij mij niet om de instance op te vragen met execQuery en zo af te sluiten
# Waarom weet ik niet
foreach (@Handles) {
      my $InParams = $Class->Methods_->{Terminate}->InParameters;
      my $relpad="Win32_Process.Handle=\"$_\"";
      my $object =  $WbemServices->get($relpad);
      print "$relpad wordt gestopt\n";
      $object->ExecMethod_("Terminate",$InParams) ;
}

# Dit werkt dus niet. Vreemde shit
# Terwijl het object wel ingeladen wordt nochtans. Dus no clue
#foreach (@Handles) {
#      my $InParams = $Class->Methods_->{Terminate}->InParameters;
#      (my $object) = $WbemServices->ExecQuery("select * where Handle=$_");
#      $object->ExecMethod_("Terminate",$InParams) ;
#}
