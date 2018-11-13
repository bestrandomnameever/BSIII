use Win32::OLE::Const;

@ARGV = ( "CDO.Message", "Excel.Sheet", "Scripting.FileSystemObject");

for $arg (@ARGV){
    $object = Win32::OLE->new($arg);

    print "############$arg###############\n";

    #constanten worden opgehaald als een ref naar een hash
    $constanten = Win32::OLE::Const->Load($object);

    # Overlopen en printen van constanten
    for $key ( sort {$a cmp $b} keys %$constanten ){
        printf("%s => %s\n", $key, $constanten->{$key});
    } 
}

