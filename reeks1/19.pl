use Win32::OLE qw(in);
use Win32::OLE::Const;

$conf=Win32::OLE->new("CDO.Configuration");
$constanten  =Win32::OLE::Const->Load($conf);

$sendMethode = $constanten->{cdoSendUsingMethod}; 
$sendPort    = $constanten->{cdoSendUsingPort};
$smtpServer  = $constanten->{cdoSMTPServer};
$conf->Fields($sendMethode)->{value} = $sendPort;
$conf->Fields($smtpServer)->{value}  = "smtp.ugent.be"; 

foreach (in $conf->{Fields}){
   print "\n\nName  : ",$_->{Name};
   print "\n\tValue : ",$_->{Value};
}
