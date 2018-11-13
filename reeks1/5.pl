use Win32::OLE qw(in);

# Alle geziene shit inladen
$excel = Win32::OLE->new("Excel.Sheet");
$fso = Win32::OLE->new("Scripting.FileSystemObject");
$cdo = Win32::OLE->new("CDO.Message");

# EnumAllObectjects returnt een callback. De sub gaat dus voor elke elemnent een functie gaan toepassen
# Hij returnt dus GEEN array!
$count = Win32::OLE->EnumAllObjects(
    sub {
        my $Object = shift;
        printf("\n%-30s : %s",ref $Object, join("/", Win32::OLE->QueryObjectType($Object)));
    }
);

print "\n$count objecten ingeladen\n";