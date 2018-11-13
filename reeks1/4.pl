use Win32::OLE qw(in with);
$cdo = Win32::OLE->new("CDO.Message");

print(ref $cdo, "/", Win32::OLE->QueryObjectType($cdo), "\n");