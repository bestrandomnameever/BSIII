use Win32::OLE qw(in);
$fso = Win32::OLE->new("Scripting.FileSystemObject");
#Krijg de huidige folder
$folder = $fso->GetFolder(".");
$files = $folder->Files();

foreach $file (in $files){
    # Hier zou eigenlijk Excel moeten staan, maar 'k heb geen Excel bestanden dus
    # PL in de plaats
    if($file->{Type} =~ /PL/){
        printf("%s\n", $file->{Name});  
    } 
}