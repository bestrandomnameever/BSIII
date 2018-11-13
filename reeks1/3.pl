use Win32::OLE qw(in with);

$cdo = Win32::OLE->new("CDO.Message");

# You can think of ref as a typeof operator.
# Hier gaat hij dus Win32::OLE printen omdat CDO.Message dit als type heeft.
print (ref $cdo, "\n");