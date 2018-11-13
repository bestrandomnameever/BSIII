use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting ';
my $DateTime = Win32::OLE->new("WbemScripting.SWbemDateTime");

sub newline {
    ( my $count ) = @_;
    $count = 1 if !defined $count;
    print "\n" x $count;
}

sub printQualifiers {
    ( my $Qualifiers ) = @_;
    for (in $Qualifiers){
        # Enige manier om te weten te komen of het een array is, is op volgende manier
        printf("%s => %s\n", $_->{Name}, ($_->{Value} =~ /ARRAY/ ? join ", ", @{$_->{Value}} : $_->{Value}) );
    }
}

my $ComputerName  = ".";
my $NameSpace = "root/cimv2";
my $WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");
my $ClassName = "Win32_Product";


$Class = $WbemServices->Get($ClassName);
$Qualifiers = $Class->{Qualifiers_};

$Instance = $WbemServices->Get($InstanceClassName);
$InstanceQualifiers = $Class->{Qualifiers_};

print "Class Qualifiers: \n";
printQualifiers($Qualifiers);
newline 2;

print "Instance Class Qualifiers: \n";
printQualifiers($InstanceQualifiers);
newline 2;






