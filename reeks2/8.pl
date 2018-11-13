use Win32::OLE qw(in with);
use Win32::OLE::Const "^Microsoft Excel";
Win32::OLE->Option(Warn => 3);

# Maak de Excel application
$excel = Win32::OLE->GetActiveObject('Excel.Application') || Win32::OLE->new('Excel.Application', 'Quit');

# Stel het pad op waar het zou moeten opgelsagen worden
$fso = Win32::OLE->new("Scripting.FileSystemObject");
$pad = $fso->GetFolder('.\Excel')->Path . "\\voud.xlsx";

# Voeg een workbook toe aan de applicatie
# We hebben dus geen idee waar deze opgeslagen zal worden op dit punt
$excel->{Workbooks}->Add();
$book = $excel->Workbooks(1);

# Om te tonen waarmee we bezig zijn
$excel->{'Visible'}=1;

# Naar de juiste sheet gaan
$sheet = $book->Worksheets(1);
$range=$sheet->Range($sheet->Cells(1,1),$sheet->Cells(49,9)); 

# De inhoud opbouwen
# Je maakt dus best een hash rij per rij, niet kolom per kolom ;)
for $i (2 .. 10){
    $teller=2;
    $inhoud->[0][$i-2] = $i;
    while($teller*$i < 100){
        $inhoud->[$teller-1][$i-2] = ($teller * $i);
        $teller++;
    }
}

# De hash in de values steken, bad practice om ze cel per cel erin te gaan rammen
$range->{Value} = $inhoud;

# Doe de volgende dingen met macro's, leer dus macro's maken in Excel:

# Eerste rij in het bold
$range->rows(1)->{'Font'}->{'Bold'} = "True";

# Lijnen tussen kolommen
$range->Borders(xlInsideVertical)->{LineStyle} = xlContinuous;
$range->Borders(xlEdgeRight)->{LineStyle} = xlContinuous;
$range->Borders(xlEdgeLeft)->{LineStyle} = xlContinuous;

# Dikke lijn onder de eerste rij
$range->rows(1)->Borders(xlEdgeBottom)->{LineStyle} = xlContinuous;
$range->rows(1)->Borders(xlEdgeBottom)->{ColorIndex} = 0;
$range->rows(1)->Borders(xlEdgeBottom)->{TintAndShade} = 0;
$range->rows(1)->Borders(xlEdgeBottom)->{Weight} = xlThick;

# Je moet blijkbaar het book saven, niet de application. Beetje vreemd
# De save method werkt ook zo
$book->SaveAs($pad);
