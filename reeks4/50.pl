use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Win32::OLE::Variant;
use Data::Dumper;

my $ComputerName  = ".";
my $NameSpace = "root/cimv2";
my $WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");

sub CheckOpFouten {
    if(Win32::OLE::LastError()){
        print Win32::OLE::LastError();
    }
}

my $NewEnvironment = $WbemServices->Get("Win32_Environment.Name='PI',UserName='<SYSTEM>'", wbemFlagUseAmendedQualifiers);
my $NewShare = $WbemServices->Get("Win32_Share.Name='Test'", wbemFlagUseAmendedQualifiers);

CheckOpFouten;
# Merk op: Delete() vs Delete_
$NewEnvironment->Delete_;
CheckOpFouten;
$NewShare->Delete();
CheckOpFouten;
