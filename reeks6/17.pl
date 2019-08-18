# Dit is de CN System
# Ik kan geen nieuwe query maken? 
# Zeker kijken dat ik verbonden ben met satan. Nu werkt dat wel gewoon hmm
# dsquery  * "DC=iii,DC=hogent,DC=be" -filter "cn=System" -s 193.190.126.71 -u "Interim F" -p "Interim F"
use Win32::OLE 'in';
use Win32::OLE::Const 'Active DS Type Library';
use Data::Dumper;
use bind_object;

my $RootDSE = bind_object::bind_object("RootDSE");
$RootDSE->{DnsHostName};
my $Object = bind_object::bind_object("CN=System" . "," . $RootDSE->{defaultNamingContext});

print $_->{cn}, "\n" foreach in $Object;
