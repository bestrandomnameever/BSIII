use Win32::OLE 'in';
use Win32::OLE::Variant;
use bind_object;
use Data::Dumper;
use Win32::OLE::Const 'Active DS Type Library';

%constanten = %{Win32::OLE::Const->Load('Active DS Type Library')};

while( (my $key, my $value) = each(%constanten) ){
    print "$key ==> $value\n" if $key =~ /.*GROUP_TYPE.*/;
}

my %choices = (1 =>  "Domain Local Distribution ", 2 => "Domain Local Security ", 3 =>  "Global Distribution",
4 => "Global Security" , 5 =>  "Universal Distribution", 6 => "Universal Security" );

# Hier gaan we de choices gaan mappen, je vind deze in de documentatie
# Ik weet niet wat zij met modulo's zit uit te steken, maar da's niet nodig
my %mappedChoices = (1 => "(groupType:1.2.840.113556.1.4.803:=" . ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP . ")"  , 2 => "(|(groupType:1.2.840.113556.1.4.803:=" . ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP . ")(groupType:1.2.840.113556.1.4.803:=" . ADS_GROUP_TYPE_SECURITY_ENABLED . "))",
    3 => "(groupType:1.2.840.113556.1.4.803:=" . ADS_GROUP_TYPE_GLOBAL_GROUP . ")",
    4 => "(|(groupType:1.2.840.113556.1.4.803:=" . ADS_GROUP_TYPE_GLOBAL_GROUP  . ")(groupType:1.2.840.113556.1.4.803:=" . ADS_GROUP_TYPE_SECURITY_ENABLED  . "))",
    5 => "(groupType:1.2.840.113556.1.4.803:=" . ADS_GROUP_TYPE_UNIVERSAL_GROUP . ")",
    6 => "(|(groupType:1.2.840.113556.1.4.803:=" . ADS_GROUP_TYPE_UNIVERSAL_GROUP  . ")(groupType:1.2.840.113556.1.4.803:=" . ADS_GROUP_TYPE_SECURITY_ENABLED  . "))"
);

# Gesorteerd schrijven van de groepen
foreach (sort {$a <=> $b} keys %choices){
    printf("%s => %s\n", $_, $choices{$_});
}


# User input vragen om de groepen te kiezen
do {
    print "Geef een nummer: ";
    chomp($nr = <>);
    push @groups, $nr if ($nr >= 1 && $nr <= 6);
} while($nr >= 1 && $nr <= 6);



# Ldap query maken
my $ADOconnection = Win32::OLE->new("ADODB.Connection");
   $ADOconnection->{Provider} = "ADsDSOObject";
   if ( uc($ENV{USERDOMAIN}) ne "III") { #niet ingelogd op het III domein
      $ADOconnection->{Properties}->{"User ID"}          = "Cedric Vanhaverbeke"; # vul in of zet in commentaar
      $ADOconnection->{Properties}->{"Password"}         = "Cedric Vanhaverbeke"; # vul in of zet in commentaar
      $ADOconnection->{Properties}->{"Encrypt Password"} = True;
   }
   $ADOconnection->Open();                                       # mag je niet vergeten
my $ADOcommand = Win32::OLE->new("ADODB.Command");
   $ADOcommand->{ActiveConnection}      = $ADOconnection;        # verwijst naar het voorgaand object
   $ADOcommand->{Properties}->{"Page Size"} = 20;

my $RootDSE=bind_object::bind_object("RootDSE"); #serverless Binding
   $RootDSE->getInfo();


my $domein = bind_object::bind_object( $RootDSE->{defaultNamingContext});

my $sBase = $domein->{adspath};

#Filter is hier een specialleke. We gaan voor elke groep gewoon een extra voorwaarde adden.

my $sFilter = "(&(objectcategory=group)";
$sFilter .= $mappedChoices{$_} foreach @groups;
$sFilter .= ")";


my $sScope      = "subtree";
my $sAttributes = "cn,groupType";

$ADOcommand->{Properties}->{"Sort On"} = "groupType";
$ADOcommand->{CommandText} = "<$sBase>;$sFilter;$sAttributes;$sScope";
my $ADOrecordset = $ADOcommand->Execute();

until ( $ADOrecordset->{EOF} ) {
    printf "%04b\t%s\n",$ADOrecordset->Fields("groupType")->{Value}   # eerste bit staat op 1
                       ,$ADOrecordset->Fields("cn")->{Value};
    $ADOrecordset->MoveNext();
}
$ADOrecordset->Close();
$ADOconnection->Close();


