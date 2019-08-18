use bind_object();
use Data::Dumper;

my $RootDSE = bind_object::bind_object("RootDSE");
print $RootDSE->get("defaultNamingContext") , "\n";

my $iii = bind_object::bind_object("OU=iii,DC=iii,DC=hogent,DC=be");

# Hier is er nu geen error normaalgezien
print "Gewoon met -> {}}";
print $iii->{"canonicalName"},"\n";
print Win32::OLE->LastError(), "\n";
print '-' x 50, "\n\n";

# Hier wel de error
print "Met Get";
print $iii->Get("canonicalName"), "\n";
print Win32::OLE->LastError(), "\n";
print '-' x 50, "\n\n";

# Met GetEx ook
print "met GetEx";
print $iii->GetEx("canonicalName"), "\n";
print Win32::OLE->LastError(), "\n";
print '-' x 50, "\n\n";.pl