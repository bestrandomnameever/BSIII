use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Data::Dumper;
my $ComputerName = ".";
my $WbemServices =  Win32::OLE->GetObject("winmgmts://$ComputerName/root/cimv2");

#bepaal alle klassen in de namespace. Je kan dit niet efficienter - een beperking op de qualifier lukt niet
my $Class = $WbemServices->Get("Win32_Directory");
foreach $Method (in $Class->Methods_){
  print Dumper $Method->InParameters;
}
