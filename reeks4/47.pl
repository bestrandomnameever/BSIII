# In WMI Reference / WMI Classes / Win32 Classes / Win32_Process staat:
# To terminate a process that you do not own, enable the SeDebugPrivilege privilege. In VBScript, you can enable this privilege with the following lines of code:
# Dus Debug instellen 

use Win32::OLE 'in';

my $ComputerName = '.';
my $Privileges="{(Debug)}!";
my $WbemServices = Win32::OLE->GetObject("winmgmts:$Privileges//$ComputerName/root/cimv2");
my $ProcessName = "notepad.exe";

$_->Terminate foreach in $WbemServices->ExecQuery("Select * From Win32_Process Where Name='$ProcessName'");