use Win32::OLE::Const;

$object = Win32::OLE->new("CDO.Message");
$conf = $object->{Configuration};

#Het kan ook zo:
#$conf = Win32::OLE->new{"CDO.Configuration"};

print "############Configuration Class###############\n";

#constanten worden opgehaald als een ref naar een hash
$constanten = Win32::OLE::Const->Load($conf);


# Overlopen en printen van constanten
for $key ( grep { $_ =~ m/SendUsing|SMTPServer/ } keys %$constanten ){
printf("%s => %s\n", $key, $constanten->{$key});
} 