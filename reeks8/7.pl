use Win32::OLE 'in';
use Win32::OLE::Variant;
use bind_object;
use Data::Dumper;
use Win32::OLE::Const 'Active DS Type Library';

my $container = bind_object::bind_object("OU=EM7INF,OU=Studenten,OU=iii,DC=iii,DC=hogent,DC=be");
$container->{Filter} = "user";

foreach(in $container){
    $_->GetInfo();

    $cn = $_->GetPropertyItem("cn", ADSTYPE_CASE_IGNORE_STRING);
    $mail = $_->GetPropertyItem("mail", ADSTYPE_CASE_IGNORE_STRING);

    printf("%s => %s\n", $cn->Values->[0]->GetObjectProperty(ADSTYPE_CASE_IGNORE_STRING), $mail->Values->[0]->GetObjectProperty(ADSTYPE_CASE_IGNORE_STRING));

    $waarde = <>;
    chomp($waarde);

    if($waarde){
        $mail->{ControlCode} =  ADS_PROPERTY_APPEND;
        $mail->{Values}->[0]->PutObjectProperty(ADSTYPE_CASE_IGNORE_STRING, $waarde);
    } else {
        $mail->{ControlCode} =  ADS_PROPERTY_CLEAR;
    }

    $_->PutPropertyItem($mail);
    $_->SetInfo();

}