use Win32::OLE qw(in);
use WIn32::OLE::Const;
use Data::Dumper;
my $Excel = Win32::OLE->new('Excel.Application');

my $constanten = Win32::OLE::Const->Load($Excel);
print Dumper $constanten;