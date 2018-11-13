use Win32::OLE 'in';
use Win32::OLE::Variant; #Nodig als je GetVarDate wil gebruiken
use Data::Dumper;

# Date ophalen om datum te kunnen parsen
my $DateTime = Win32::OLE->new("WbemScripting.SWbemDateTime");

# Informatie
$ComputerName = ".";
$NameSpace = "root/cimv2";
$ClassName = "Win32_Environment.Name=\"TEST\",UserName=\"<SYSTEM>\"";

# WbemServices maken
$monik = "winmgmts://$ComputerName/$NameSpace";
$WbemServices = Win32::OLE->GetObject($monik);

$instances = $WbemServices->ExecQuery("select * from Win32_NTLogEvent where EventCode = 6005 or EventCode = 6006");

# In de oplossing staat dat ze sorteren op TimeWritten, maar dat kan ook RecordNumber zijn
for (sort {$a->{RecordNumber} cmp $b->{RecordNumber}} in $instances){

    $DateTime->{Value} = $_->{TimeWritten};
    $periode = ($DateTime->GetFileTime - $StartupTime)/10000000; #Waarom dat delen hier? Blijkbaar staat dit in iets zeer groot ofzo?
    printf("%-22s: %s %s\n", $DateTime->GetVarDate,
                                $_->{Message} =~/(started|stopped)/, #Deze regex geeft ofwel started of stopped terug
                                ($_->{EventCode} == 6006 && $StartupTime ? "na $periode s" : "")); #Deze gaat alleen schrijven als EventCode 6006 (stopped) is net na een started want anders is StartupTime undef
    #$StartupTime = ($_->{EventCode} == 6005 ? $DateTime->GetFileTime : undef);
    $StartupTime=($_->{EventCode} == 6005 ? $DateTime->GetFileTime : undef);
}
