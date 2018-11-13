use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Win32::OLE::Variant;
use Data::Dumper;

my $DateTime = Win32::OLE->new("WbemScripting.SWbemDateTime");

@ARGV = ("Win32_Group", "Win32_Process", "Win32_Volume", "Win32_Share");

my $ClassName = shift;

my $ComputerName  = ".";
my $NameSpace = "root/cimv2";
my $WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");

sub printFormattedHash {
    my $hash = shift;
    my %eenhash = ();
    my %tweehash = ();

    foreach $Method (@{$hash->{1}}){
        print $Method->{Name}, " Met een ValueMap en Values\n";
        print '-' x 50;
        print "\n";
        @eenhash{ @{$Method->Qualifiers_->Item("ValueMap")->{Value}} } = @{$Method->Qualifiers_->Item("Values")->{Value}};
        while( ($key, $value) = each(%eenhash) ){
            printf("%s => %s\n", $key, $value);
        }
        print "\n";
    }

    foreach $Method (@{$hash->{2}}){
        print $Method->{Name}, " Met enkel Values\n";
        print '-' x 50;
        print "\n";
        @array = @{$Method->Qualifiers_->Item("Values")->{Value}};
        for (0 .. scalar @array - 1){
            $tweehash{$_} = $array[$_];
        }
        
        while( ($key, $value) = each(%tweehash) ){
            printf("%s => %s\n", $key, $value);
        }
        print "\n";
    }
        print "\n\n";
}

sub getMethodsWithTextReturnValue {
    my $Class = shift;
    my @Methods;
    # in 1 steek ik de methoden die een valuemap en values hebben
    # in 2 steek ik degenen die enkel een values attribuut hebben
    for my $Method (in $Class->Methods_){
        push @{$hash->{1}}, $Method if $Method->Qualifiers_->Item("ValueMap") && $Method->Qualifiers_->Item("Values");
        push @{$hash->{2}}, $Method if !$Method->Qualifiers_->Item("ValueMap") && $Method->Qualifiers_->Item("Values");
    }
    return $hash;
}

while(scalar @ARGV != 0){
    my $Class = $WbemServices->Get(shift, wbemFlagUseAmendedQualifiers);
    printf("%s\n", $Class->Path_->{Class});
    print '=' x 70;
    print "\n\n";
    $hash = getMethodsWithTextReturnValue($Class);
    printFormattedHash($hash);

}

