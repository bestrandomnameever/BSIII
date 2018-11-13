use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Win32::OLE::Variant;
use Data::Dumper;

my $DateTime = Win32::OLE->new("WbemScripting.SWbemDateTime");

my $ComputerName  = ".";
my $NameSpace = "root/cimv2";
my $ClassName = "Win32_Process";

my $WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");

sub PrintMethodQualifiers {
    my $Method = shift;
    for $Qualifier (in $Method->Qualifiers_){
        
        printf("%s => %s\n", $Qualifier->{Name}, ($Qualifier->{Value} =~ /array/i ? join ", ", @{$Qualifier->{Value}} : $Qualifier->{Value}));

    }
}

sub geefOverzichtVanMethoden{
    (my $object) = shift;
    printf("%s\n", $object->Path_->{Class});
    print '=' x 50;
    print "\nHeeft " . $object->Methods_->Count . " methoden: \n\n\n";
    $Methods = $object->{Methods_};
    for (in $Methods){
        printf("%s\n",$_->{Name});
        print '-' x 50;
        print "\n" . $_->Qualifiers_->Item(Description);
        PrintMethodQualifiers($_);
        print "\n";
    }

}
$Class = $WbemServices->Get($ClassName, wbemFlagUseAmendedQualifiers);
geefOverzichtVanMethoden($Class);