# De klasse Win32_OperatingSystem heeft die info
use Win32::OLE 'in';
use Win32::OLE::Const;
Win32::OLE->Option(Warn => 3);

# Values voor CIMType te weten komen, anders gaan er gewoon nummers staan bij CIMType en dat willen we niet.
my $Locator=Win32::OLE->new("WbemScripting.SWbemLocator");

#Direct casten naar een hash
%wd = %{Win32::OLE::Const->Load($Locator)}; 
my %types;

while (($type,$nr)=each (%wd)){
  if ($type=~/Cimtype/){
    $type=~s/wbemCimtype//g;
    $types{$nr}=$type;
  }
}

# Nu alles uitschrijven

$ComputerName = ".";
$NameSpace = "root/cimv2";

$ClassName = "Win32_Directory";

$moniker = "winmgmts://$ComputerName/$NameSpace:$ClassName";
$Directory = Win32::OLE->GetObject($moniker);

foreach $property (in $Directory->Properties_, $Directory->SystemProperties_){
        printf("%s type: %s", $property->{Name}, $types{$property->{CIMType}});
        print $property->{Isarray} ? " - is een array\n" : "\n";  
}