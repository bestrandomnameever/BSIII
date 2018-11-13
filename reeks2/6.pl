use Win32::OLE qw(in with);
use Win32::OLE::Const "^Microsoft Excel";
Win32::OLE->Option(Warn => 3);

sub printRange {
    ($cells) = @_;
    $cells = $range->{'Value'};

    foreach $ref_cells (@$cells){
        foreach $scalar (@$ref_cells){
            printf("%10s", $scalar);
        }
        print "\n";
    }
}

@ARGV = ( "Book1.xlsx" );
$fso = Win32::OLE->new("Scripting.FileSystemObject");
$Excel = Win32::OLE->GetActiveObject('Excel.Application') || Win32::OLE->new('Excel.Application', 'Quit');
$Excel->{'Visible'} = 1;

$Book = $Excel->Workbooks->Open(join('\\', $fso->GetFolder('.\Excel')->Path, $ARGV[0]));
$sheet = $Book->Worksheets(1);
$range=$sheet->Range("A1:D10");
@_ = ($range);
printRange();
$range=$sheet->Cells(4,1);
@_ = ($range);
printRange();
$range=$sheet->Range($sheet->Cells(1,1),$sheet->Cells(4,3));
@_ = ($range);
printRange();