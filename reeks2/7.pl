use Win32::OLE qw(in with);
use Win32::OLE::Const "^Microsoft Excel";
Win32::OLE->Option(Warn => 3);

@ARGV = ( "Book1.xlsx" );
$fso = Win32::OLE->new("Scripting.FileSystemObject");
$Excel = Win32::OLE->GetActiveObject('Excel.Application') || Win32::OLE->new('Excel.Application', 'Quit');

$Book = $Excel->Workbooks->Open(join('\\', $fso->GetFolder('.\Excel')->Path, $ARGV[0]));
$sheet = $Book->Worksheets(1);

# Range maken 
$range = $sheet->Range("A2:C7");

# Range vullen
$range->{Value} = [
    ["Delivered", "En Route", "To be shipped"],
    [504, 102, 86],
    [670, 150, 174],
    [891, 261, 201],
    [1274, 471, 321],
    [1563, 536, 241]
];

# Savet de inhoud van de workspace
# Je kan deze methode vinden in MSDN/Office Solutions Development/2007 Microsoft Office System/Excel 2007/Excel Object Model Reference/Application Object/Application Object Members
# Hier moet je zoeken naar deze methode.

# Een andere methode tonen ze in de oplossing, daar save je op je workbook
$Excel->SaveWorkspace();
