use bind_object;
use Data::Dumper;

$RootObj = bind_object::bind_object("RootDSE");
my $cont = bind_object::bind_object("OU=Studenten,OU=iii,".$RootObj->get(defaultNamingContext));

# Hier nog geen GetInfo gedaan dus 0
print $cont->{PropertyCount}, " properties in Property Cache\n";
$cont->getInfo();

# Hier wel al GetInfo gedaan dus alles ingeladen behalve de geconstrueerde
print $cont->{PropertyCount}, " properties in Property Cache\n";

# Hier de specifieke in de lijst ingelade (dit zijn geconstrueerde) en degene die er al in zaten
$cont->getInfoEx(["ou","canonicalName","msDS-Approx-Immed-Subordinates","objectclass"],0);
print $cont->{PropertyCount}, " properties in Property Cache\n";