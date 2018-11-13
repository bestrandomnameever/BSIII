use Win32::OLE;

$bestandsnaam = $ARGV[0];
$fso = Win32::OLE->new("Scripting.FileSystemObject");
if($fso->FileExists($bestandsnaam)){
    print "Bestand bestaat\n";

    # Returnt een file object
    # Info over fileobject kan je vinden door door te klikken naar 
    # fileSystemObject Reference > Object > File Object > Properties

    $file = $fso->GetFile($bestandsnaam);
    printf("Absoluut pad op 2 manieren:\n%s\n%s\n", $fso->GetAbsolutePathName($bestandsnaam), $file->{Path});
    printf("Type van het bestand: %s\n", $file->{Type});
} else {
    print "Bestand bestaat niet\n";
} 

