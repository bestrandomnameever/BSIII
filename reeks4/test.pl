use Win32::OLE qw(in);
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Data::Dumper;

my $Service = Win32::OLE->GetObject('winmgmts://./root/cimv2');
my $Class = $Service->Get('Win32_Process');
my $Method = $Class->Methods_->{Create};
my $Privilege = $Method->Qualifiers_->{Privileges};

print join ",", map {/Se(.*)Privilege/} @{$Privilege->Value};
