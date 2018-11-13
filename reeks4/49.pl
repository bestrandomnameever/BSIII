use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Win32::OLE::Variant;
use Data::Dumper;

my $ComputerName  = ".";
my $NameSpace = "root/cimv2";
my $WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");

my $ClassName = "Win32_Environment";
my $class = $WbemServices->Get($ClassName);

sub CheckOpFouten {
    if(Win32::OLE::LastError()){
        print Win32::OLE::LastError();
        return 1;
    }
    return 0;
}

my $instance = $class->SpawnInstance_();
# Ik vul nu de sleutelattributen in van deze nieuwgemaakte klasse: 
# Ook de properties die overgeërfd worden zijn interessant
$instance->{UserName} = '<SYSTEM>';
$instance->{Name} = "PI";
$instance->{VariableValue} = "3.1415";
$instance->{SystemVariable} = 1;

my $Path = $instance->Put_();

# Variabele is gemaakt
if(!CheckOpFouten){
    print "Object gemaakt!\n";
}

print "Absolute path = ",$instance->{path_}->{path},"\n"; #geen uitvoer
print "Absolute path = ",$instance->{SystemProperties_}->Item("__PATH")->{Value},"\n"; #geen uitvoer
#lukt wel
print "return        = ",$Path->{path},"\n";     

#relatieve pad lukt op drie manieren
print "Relative path = ",$instance->{path_}->{relpath},"\n";                             
print "Relative path = ",$instance->{SystemProperties_}->Item("__RELPATH")->{Value},"\n";
print "return        = ",$Path->{RelPath},"\n";
