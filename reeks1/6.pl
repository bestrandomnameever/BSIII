use Win32::OLE qw(in with);

$object = Win32::OLE->new("fault");

# Hier komt de laatste fout in
printf("Laatste fout: %s\n", Win32::OLE->LastError());