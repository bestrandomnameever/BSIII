use Win32::OLE;
use bind_object;
use Data::Dumper;

# We moeten in het schema gaan kijken
my $ADOconnection = Win32::OLE->new("ADODB.Connection");
$ADOconnection->{Provider} = "ADsDSOObject";

if ( uc($ENV{USERDOMAIN}) ne "III") { #niet ingelogd op het III domein
    $ADOconnection->{Properties}->{"User ID"}          = "Cedric Vanhaverbeke"; # vul in of zet in commentaar op school
    $ADOconnection->{Properties}->{"Password"}         = "Cedric Vanhaverbeke"; # vul in of zet in commentaar op school
    $ADOconnection->{Properties}->{"Encrypt Password"} = True;
}

$ADOconnection->Open();

my $ADOcommand = Win32::OLE->new("ADODB.Command");
$ADOcommand->{ActiveConnection} = $ADOconnection;        # verwijst naar het voorgaand object
$ADOcommand->{Properties}->{"Page Size"} = 20;     

Win32::OLE->LastError()&& die (Win32::OLE->LastError()); # Is er een fout gebeurtd?

my $RootDSE = bind_object::bind_object("RootDSE");
$RootDSE->GetInfo();
my $Schema = bind_object::bind_object($RootDSE->{schemaNamingContext});

# Enkel klassen gaan ophalen, je kan dat eventueel ook straks pas doen, maar zo ging dat in de vorige reeks
#$Schema->{Filter} = "classSchema";

my $sBase = "<$Schema->{ADsPath}>";

my $sFilter     = "(&(objectclass=classSchema)(!(objectClassCategory=1)))";
my $sAttr = "cn,objectClassCategory";
my $sScope      = "subtree";

$ADOcommand->{CommandText} = "$sBase;$sFilter;$sAttr;$sScope";

my $ADOResultSet = $ADOcommand->Execute();
Win32::OLE->LastError()&& die (Win32::OLE->LastError()); # Is er een fout gebeurtd?

# In de oplossing overlopen ze ADOResultSet twee keer en doen ze 
# $ADOResultSet->MoveFirst() zodat terug naar de eerste entry gekeken wordt. Dat lijkt mij hier zinloos maar soit

while(!$ADOResultSet->{EOF}){
    push @Auxiliary, $ADOResultSet->Fields("cn")->{Value} if $ADOResultSet->Fields("objectClassCategory")->{Value} == 2;
    push @Abstract, $ADOResultSet->Fields("cn")->{Value} if $ADOResultSet->Fields("objectClassCategory")->{Value} == 3;
    $ADOResultSet->MoveNext();
}

$ADOResultSet->Close();
$ADOconnection->Close();

print "Hulpklassen:", "\n";
print "-" x 50, "\n";
foreach (sort {lc($x) cmp lc($y)} @Auxiliary){
    print $_,"\n";
}

print "\n" x 3;

print "Abstracte klassen:\n";
print "-" x 50, "\n";
foreach (sort {lc($x) cmp lc($y)} @Abstract){
    print $_,"\n";
}

