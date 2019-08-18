use Win32::OLE;
use bind_object;

my $RootDSE = bind_object::bind_object("RootDSE");
my $DefaultSchema = $RootDSE->{defaultNamingContext};

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
$ADOcommand->{Properties}->{"Sort On"} = "printerName";

# Nu alle delen van het command doen
# We doen het in LDAP syntax
my $Where = $DefaultSchema->{ADsPath}; # AdsPath van het schema is de root
my $Criteria = "(&(objectCategroy=printQueue)(printColor=TRUE)(printDuplexSupported=TRUE))";
my $sAttributes = "printRate,printMaxXExtent,whenChanged,printStaplingSupported,objectClass,printMaxResolutionSupported,printername,adspath,objectGUID";
my $Scope = "subtree";

$ADSCommand->{CommandText} = "<$Where>;$Criteria;$sAttributes;$Scope";
