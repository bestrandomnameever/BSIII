use Win32::OLE qw(in);


$message = Win32::OLE->new("CDO.Message");
Win32::OLE->Option(Warn => 3);

# Configuratie van het bericht ophalen
$conf = $message->{Configuration};

# Je vindt de properties in de Imessage interface
# afzender, de bestemmeling, het onderwerp en de inhoud instellen
$message->{Sender} = "cedric.vanhaverbeke\@ugent.be";
$message->{To} = "cedric.vanhaverbeke\@hotmail.com";
$message->{Subject} = "Test";

# Ik wou dit eens testen. Je zou wellicht ook gewoon de TextBody kunnen instellen
$message->{HTMLBody} = "<h1>Dit is een test</h1>";
$message->{AutoGenerateTextBody} = "true";

# De configuratie van het bericht instellen
$conf->Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver")->{Value}     = "smtp.ugent.be" #zie opmerkingen bij oef 13
$conf->Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport")->{Value} = 25;              #niet altijd noodzakelijk
$conf->Fields("http://schemas.microsoft.com/cdo/configuration/sendusing")->{Value}      = 2;
$conf->{Fields}->Update();      #is noodzakelijk

foreach (in $conf->{Fields}){
   printf "%s = %s\n",$_->{Name},$_->{Value};
}

# Verzenden van bericht
$message->Send();
