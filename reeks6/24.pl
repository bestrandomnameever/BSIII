use bind_object;
use Data::Dumper;

# Volgens mij:
#Getinfo
#Setinfo
#Get
#Put
#GetEx
#PutEx
#GetInfoEx

my $RootDSE = bind_object::bind_object("RootDSE");
$RootDSE->GetInfo();
print $RootDSE->{defaultNamingContext} , "\n";

# OF met get zonder getInfo;

