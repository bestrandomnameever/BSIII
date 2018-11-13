#Dit script haalt gewoon alle used cellen op in elke worksheet

use Win32::OLE qw(in with);
@ARGV = ( "Book1.xlsx" );
$fso = Win32::OLE->new("Scripting.FileSystemObject");
$Excel = Win32::OLE->GetActiveObject('Excel.Application') || Win32::OLE->new('Excel.Application', 'Quit');
$Excel->{'Visible'} = 1;

$Book = $Excel->Workbooks->Open(join('\\', $fso->GetFolder('.\Excel')->Path, $ARGV[0]));

for (1 .. $Book->Worksheets->Count){
    $cellen =  $Book->Worksheets($_)->UsedRange;
    printf("Aantal cellen in Worksheet %d: %d\n", $_, $cellen->Count);
}

