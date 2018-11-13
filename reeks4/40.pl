use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Win32::OLE::Variant;
use Data::Dumper;

my $DateTime = Win32::OLE->new("WbemScripting.SWbemDateTime");

my $ComputerName  = ".";
my $NameSpace = "root/cimv2";
my $WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");

my $ClassName =  "Win32_Share";
my $Class = $WbemServices->Get($ClassName, wbemFlagUseAmendedQualifiers);

sub telOutputParameters {
    my $Method = shift;
    return 
}




for my $Method (in $Class->Methods_){
    print $Method->Name . "\n";

    #INPUTPARAMS
    printf("%s \n", "INPUTPARAMETERS");
    if($Method->InParameters){
        for my $property (in $Method->InParameters->Properties_){
                # Vragen of je ergens grafisch kan zien wat inputparameters zijn van een methode
                printf("InputParameter: %s\n", ($property->Qualifiers_->{Optional} ? "[" . $property->{Name} . "]" : $property->{Name}));
        }
    } else {
        print "No input parameters\n";
    }

    #OUTPUTPARAMS
    printf("%s \n", "OUTPUTPARAMETERS");
    if($Method->OutParameters && $Method->OutParameters->Properties_->Count > 1){
        for my $property (in $Method->OutParameters->Properties_){
                printf("OutputParameter: %s\n", $property->{Name});
        }
    } else {
        print "No more than 1 outputparameter";
    }

    #CHECK IF STATIC
    printf("Static method?: %s\n", ($Method->Qualifiers_->Item("Static")) ? "true" : "false" );
    print "\n\n";
}