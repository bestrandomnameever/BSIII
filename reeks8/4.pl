use bind_object;
use Win32::OLE qw(in);
use Data::Dumper;
use Win32::OLE::Variant;
use Win32::OLE::Const "Active DS Type Library";

# Dit staat hier om random names voor organizational units te genereren.
my @chars = ("A" .. "Z", "a" .. "z");
my $string;
$string .= $chars[rand @chars] for 1..8;

@ARGV = ($string);
$RootDSE = bind_object::bind_object("RootDSE");


@ARGV = ("Cedric Vanhaverbeke");

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
    #$names{(split('/',$_->Get("canonicalName")))[-1]} = undef;
    $names{$_->Get("cn")} = undef;
}


die "Naam bevindt zich al in de container\n" if HashContainsName(\%names, $ARGV[0]);

my $newOrganizationalUnit = $container->Create("organizationalUnit", "ou=$ARGV[0]");

die "Fout bij het maken van de container\n" if !$newOrganizationalUnit;

# Wijzigingen doorvoeren
#$newOrganizationalUnit->SetInfo();

print "Toegevoegd\n" if !Win32::OLE->LastError();


