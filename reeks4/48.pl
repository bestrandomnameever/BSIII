use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Win32::OLE::Variant;
use Data::Dumper;

my $ComputerName  = ".";
my $NameSpace = "root/cimv2";
my $WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");

my $ClassName = "Win32_Share";
my $class = $WbemServices->Get($ClassName, wbemFlagUseAmendedQualifiers);

sub CheckOpFouten {
    if(Win32::OLE::LastError()){
        print Win32::OLE::LastError();
    }
}

# In WMI gekeken wat de createBy is en of SupportsCreate op true staat.
# CreateBy = Create. Dus deze methode gebruiken

# Bij method parameters kan je de parameters zien
# 1. Path (moet blijkbaar een bestaand path zijn)
# 2. Name
# 3. Type (Moet een int zijn, geen idee waarvoor alles staat)
# 4. MaximumAllowed [Optional]
# 5. Description [Optional]
# 6. Password [Optional]
# 7. Access [Optional]

# Directe methode
$ReturnValue = $class->Create("C:\\temp", "Test", 0);
print $ReturnValue; # Er is geen tekstuele interpretatie

# Checken of er iets fout gaat bij het creeëren

# Formele methode, daarmee bedoelen ze niet de statische methode

my $InParams = $class->Methods_->{Create}->InParameters;

$InParams->{Path} = "C:\\temp";
$InParams->{Name} = "test";
$InParams->{Type} = 0;

my $ReturnValue = $class->ExecMethod_("Create", $InParams);
print $ReturnValue->{ReturnValue};

# Als je wil zien wat er juist in de $ReturnValue zit met ExecMethod kan je alle properties van de $ReturnValue overlopen op volgende manier
#foreach (in $ReturnValue){
#    print Dumper $_;
#}




