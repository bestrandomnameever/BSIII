use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting ';
my $DateTime = Win32::OLE->new("WbemScripting.SWbemDateTime");

sub newline {
    ( my $count ) = @_;
    $count = 1 if !defined $count;
    print "\n" x $count;
}

sub checkSingleton {
    ( my $Qualifiers )  = @_;
    for (in $Qualifiers){
        if($_->{Name} =~ /Singleton/){
            return 1;
        }
    }
    return 0;
}

my $ComputerName  = ".";
my $NameSpace = "root/cimv2";
my $WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");
my $ClassName = "Win32_OperatingSystem=@";

$Class = $WbemServices->Get($ClassName);
$Qualifiers = $Class->{Qualifiers_};

$isSingleton = checkSingleton($Qualifiers);

printf("%s", $isSingleton ? "Klasse is een singelton" : "Klasse is geen singleton");
newline;






