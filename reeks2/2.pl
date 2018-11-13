use Win32::OLE qw(in with);

$fso = Win32::OLE->new("Scripting.FileSystemObject");

@ARGV = ( "Book1.xlsx" );

$Excel = Win32::OLE->GetActiveObject('Excel.Application') || Win32::OLE->new('Excel.Application', 'Quit');
$Excel->{'Visible'} = 1;

$Book = $Excel->{Workbooks}->Open(join('\\', $fso->GetFolder('.\Excel')->{Path}, $ARGV[0]));
$array = $Book->{Worksheets};
printf("In deze workbook zitten %d sheets\n",$array->{Count});
