use Win32::OLE qw(in with);
use Data::Dumper;

$fso = Win32::OLE->new("Scripting.FileSystemObject");
$excel = Win32::OLE->GetActiveObject('Excel.Application') || Win32::OLE->new('Excel.Application', 'Quit');

$naam = shift;
if($fso->FileExists($naam)){
    $Book = $Excel->{Workbooks}->Open(join('\\', $fso->GetFolder('.\Excel')->{Path}, $ARGV[0]));
}

print Win32::OLE->LastError();