use Win32::OLE 'in';
use Win32::OLE::Variant;
my $DateTime = Win32::OLE->new("WbemScripting.SWbemDateTime");

my $ComputerName  = ".";
my $NameSpace = "root/cimv2";
my $WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");
my $ProcessName="Excel.exe";
my $QueryProcess = "Select * From Win32_Process Where Name='$ProcessName'";

my $NameSpaceMS = "root/msapps12";
my $WbemServicesMS = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpaceMS"); 
my $QueryExcel = "Select * From Win32_ExcelWorkBook";

sub printproperties{
    #Object ophalen om properties van te printen
    (my $object) = shift;

    foreach (in $object->Properties_, $object->SystemProperties){
        # Als het een array is, print dan de array. Anders gewoon de value
        schrijf($_->{Name}, ($_->{IsArray} ? join ",", @{$_->{Value}} : $_->{Value} , $_->{CimType}) );
    }
}

sub schrijf {
    # Deconstruct
    ( my $Name, my $Value, my $Type ) = @_;
    if($Type != 101){
        printf("\t%20s =%s (%s)\n", $Name, $Value, $Type);
    } else {
        $DateTime->{Value} = $Value;
        printf("\t%20s =%s (%s)\n", $Name, $DateTime->GetVarDate, $Type);
    }
}

sub geefstatus{
    (my $info) = shift; # Haalt het eerste argument van @ARGV. Is hetzelfde als $ARGV[0] schrijven
    printf("--------------------- %s -----------------\n", $info);

    # Nu voor elk proces dat te maken heeft met Excel alle properties uitschrijven
    foreach (in $WbemServices->ExecQuery($QueryProcess)){
        printproperties($_);
    }

    # Hier voor alle Microsoft manipulaties in Excel
    foreach (in $WbemServicesMS->ExecQuery($QueryExcel)){
        printproperties($_);
    }

    print "\n";
}

my $Excel =  Win32::OLE->new('Excel.Application', 'Quit');
geefstatus("Nieuwe Excel Applicatie gestart");
my $Excel = Win32::OLE->GetActiveObject('Excel.Application');
geefstatus("Draaiende Excel-applicatie gebruiken");
my $Book = $Excel->Workbooks->Add();
geefstatus("Werkbook toegevoegd");
$Book->Save();
geefstatus("Werkboek opgeslaan");
print "\n\n";