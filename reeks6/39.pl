# Toegang gebeurt via de ldapDisplayName, niet de distinguished name

use bind_object;
use Data::Dumper;
use Win32::OLE::Variant;

my $RootDSE = bind_object::bind_object("RootDSE");
$Schema = bind_object::bind_object($RootDSE->get(schemaNamingContext));

for(("OrganizationalUnit", "cn", "printerName")){
@ARGV = ($_);

my $Object = bind_object::bind_object("schema/$ARGV[0]");

#GUID is dezelfde voor alle abstracte klassen en abstracte attributen
# Dit zijn de 6 attributen van de IADs Interface
foreach ("adspath","class","GUID","name","parent","Schema"){ 
  printf "%20s : %s\n",$_,$Object->{$_};
}

#vs 

# Werkt niet
$Object->GetInfoEx(["AdsPath","Class","GUID","Name","Parent","Schema"], 0);
foreach ("AdsPath","Class","GUID","Name","Parent","Schema"){ 
  printf "%20s : %s\n",$_,$Object->get($_);
  print Win32::OLE::LastError();
}

# Je kan zo dus de canonical name opvragen van het reeël object
# Wat je daarmee precies bent is mij ook een raadsel
print "Het overeenkomstig reeel object heeft cn=",$Object->get("cn");
print "\n\n";
}
