use Win32::OLE 'in';
use Win32::OLE::Variant;
use bind_object;
use Data::Dumper;
use Win32::OLE::Const 'Active DS Type Library';

@ARGV = ($string);
$RootDSE = bind_object::bind_object("RootDSE");

@ARGV = ("Test");

my $container = bind_object::bind_object("OU=EM7INF,OU=Studenten,OU=iii," . $RootDSE->Get("defaultNamingContext"));

sub HashContainsName {
    (my $hashref, my $name) = @_;
    foreach(keys %{$hashref}){
        if(lc($_) eq lc($name)){
            return 1;
        }
    }
    return 0;
}

# De naam van de organizational unit moet altijd uniek zijn, daarom gaan we eerst de naam van de andere ophalen
foreach (in $container){
    $_->GetInfoEx(["canonicalName"], 0);
    $names{(split('/',$_->Get("canonicalName")))[-1]} = undef;
}


die "Naam bevindt zich al in de container\n" if HashContainsName(\%names, $ARGV[0]);

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


my $domein = bind_object::bind_object( $RootDSE->{defaultNamingContext});
my $sBase = $domein->{adspath};

#Filter is hier een specialleke. We gaan voor elke groep gewoon een extra voorwaarde adden.

my $sFilter = "(&(objectcategory=group)(samAccountName=$ARGV[0]))";
my $sScope      = "subtree";
my $sAttributes = "samAccountName";

$ADOcommand->{Properties}->{"Sort On"} = "groupType";
$ADOcommand->{CommandText} = "<$sBase>;$sFilter;$sAttributes;$sScope";
my $ADOrecordset = $ADOcommand->Execute();

die "samAccountName bestaat al" if !$ADOrecordset_>{EOF};
$ADOrecordset->Close();
$ADOconnection->Close();

$newGroup = $container->Create("group", "cn=$ARGV[0]");
die "Niet gelukt om groep aan te maken" if Win32::OLE->LastError();

$newGroup->Put("samAccountName", $ARGV[0]);
$newGroup->SetInfo();
print "Groep aanmaken gesimuleerd!\n" if Win32::OLE->LastError();

$newGroup->GetInfo();


$schema = bind_object::bind_object($newGroup->{Schema});

foreach my $property (@{$schema->{MandatoryProperties}}){
    printf("%s => %s\n", $property, $newGroup->{$property});
}

