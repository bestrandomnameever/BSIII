use Win32::OLE qw(in);
use Win32::OLE::Const;

$conf=Win32::OLE->new("CDO.Configuration");
$message=WIN32::OLE->{"CDO.Message"};
$constanten  =Win32::OLE::Const->Load($conf);

$sendMethode = $constanten->{cdoSendUsingMethod}; 
$sendPort    = $constanten->{cdoSendUsingPort};
$smtpServer  = $constanten->{cdoSMTPServer};

# Configuratie heeft altijd een attribuut "Fields"
# Je kan dus niet rechtstreeks de Fields gaan opvragen
$conf->Fields($sendMethode)->{value} = $sendPort;
$conf->Fields($smtpServer)->{value}  = "smtp.ugent.be"; 

# Configuratie instellen
$message->{Configuratie} = $conf;
$message->Send();

