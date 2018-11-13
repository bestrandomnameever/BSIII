use Win32::OLE;
use Data::Dumper;
use Win32::OLE 'in';

#my $shell=Win32::OLE->new("WScript.Shell");
#$shell->LogEvent(2,"Schrijf message in logboek");

# In het logboek 'Application' zou nu iets moeten bijgekomen zijn
# logboek is van het type Win32_NTEventlogFile

# Informatie
$ComputerName = ".";
$NameSpace = "root/cimv2";


# WbemServices maken
$monik = "winmgmts://$ComputerName/$NameSpace";
$WbemServices = Win32::OLE->GetObject($monik);

# LogApplication ophalen
my $logboek="Application";
my $sourceName="WSH";  #kan je opzoeken in het logboek.
# Zoek hiervoor in het logboek naar de meest recente logfile, daar vind je al zijn properties
my $Query="Select * from Win32_NTLogEvent where Logfile='$logboek' and SourceName='$sourceName'";

my  $EventInstances = $WbemServices->ExecQuery($Query);

printf "\n%s events gevonden in %s\n",$EventInstances->{Count},$logboek;
foreach (   in $EventInstances) {
    #$tekst =$_->{InsertionStrings};

    #printf "\n\t Inhoud:%s",join("\n",@{$tekst}); #Dit vertaalt de array die ref in het object naar een perl-array
    # Volgens mij werkt dit ook gewoon:
    printf("Inhoud: %s\n", $_->{Message})
}
