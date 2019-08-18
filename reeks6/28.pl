use bind_object;
use Data::Dumper;

my $Object = bind_object::bind_object("OU=Studenten,OU=iii,DC=iii,DC=hogent,DC=be");

# Zodat enkel de organizational units getoond worden
$Object->{Filter} = "organizationalUnit";
foreach(in $Object){
    # Omdat het een geconstrueerd attribuut is moet je eerst de info gaan ophalen 
    $_->GetInfoEx( [ "ou", "msDS-Approx-Immed-Subordinates" ], 0);
    printf("%s: %d\n", $_->Get("ou"), $_->Get("msDS-Approx-Immed-Subordinates"));
}