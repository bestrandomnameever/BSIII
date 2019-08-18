use Win32::OLE;


# Connection maken
my $ADSConnection = Win32::OLE->new("ADODB.Connection");
$ADSConnection->{Provider} = "ADsDSOObject";

# Voor als je thuis bent
$ADSConnection->{Properties}{"User ID"} = "Cedric Vanhaverbeke";
$ADSConnection->{Properties}{"Password"} = "Cedric Vanhaverbeke";
$ADSConnection->{Properties}{"Encrypt Password"} = True;

# Niet vergeten openen
$ADSConnection->Open();

# Command linken met de connectie
my $ADSCommand = Win32::OLE->new("ADODB.Command");
$ADSCommand->{AciveConnection} = $ADSConnection;
$ADSCommand->{Properties}{"Page Size"} = 20;
$ADSCommand->{CommandText} = "SELECT AdsPath, cn FROM 'LDAP://DC=Fabrikam,DC=com' WHERE objectClass = '*'";