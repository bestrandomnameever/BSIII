use Win32::OLE qw(in with);
use Win32::OLE::Const "^Microsoft Excel";
Win32::OLE->Option(Warn => 3);

@ARGV = ( "Book1.xlsx" );
$fso = Win32::OLE->new("Scripting.FileSystemObject");
$Excel = Win32::OLE->GetActiveObject('Excel.Application') || Win32::OLE->new('Excel.Application', 'Quit');
$Excel->{'Visible'} = 1;

$Book = $Excel->Workbooks->Open(join('\\', $fso->GetFolder('.\Excel')->Path, $ARGV[0]));

for $nsheet (in $Book->Worksheets){
    $laatstegebruikteCell = $nsheet->Range('A1')->SpecialCells(xlCellTypeLastCell);
    $range = $nsheet->Range('A1', $laatstegebruikteCell);

    $cells = $range->{Value};

    foreach $ref_cells (@$cells){
        foreach $scalar (@$ref_cells){
            printf("%10s", $scalar);
        }
        print "\n";
    }

}