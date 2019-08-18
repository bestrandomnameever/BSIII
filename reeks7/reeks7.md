# Reeks 7

## dsquery

Je kan makkelijk queries controleren met dsquery vooraleer je ze gaat schrijven.
TODO: uitleg dsquery

## Een ldap query maken

Een ldap query maken gaat altijd op dezelfde manier:

1. Een ADODBConnection maken:

   - Een provider meegeven
   - Eventuele properties meegeven
   - de connectie openen

2. Een ADODBCommand maken

   - De connectie meegeven
   - Het filterpad maken
   - Het commando uitvoeren en dit gelijk stellen aan een ResultSet

3. De resultset overlopen

   - Zolang we niet op EOF zitten
   - Attributen ophalen
   - MoveNext()

4. Alles closen

```perl
my $RootDSE = bind_object::bind_object("RootDSE");

# Connection
my $ADOConnection = Win32::OLE->new("ADODB.Connection");
$ADOConnection->{Provider} = "ADsDSOObject";
if(uc($ENV{USERDOMAIN} ne 'III')){
    $ADOConnection->{Properties}{"User ID"} = "Interim F";
    $ADOConnection->{Properties}{"Password"} = "Interim F";
    $ADOConnection->{Properties}{"Encrypt Password"} = True;
}

$ADOConnection->Open();

# Command
my $ADOCommand = Win32::OLE->new("ADODB.Command");
$ADOCommand->{ActiveConnection} = $ADOConnection;
$ADoCommand->{Properties}{"Page Size"} = 20;


my $domein = bind_object::bind_object("OU=EM7INF,OU=Studenten,OU=iii," . $RootDSE->Get("defaultNamingContext"))->{AdsPath};

my $sBase = $domein;
my $sFilter = "(&(objectCategory=person)(objectClass=user))";
my $sAttributes = "distinguishedName";
my $sScope = "subtree";

$ADOCommand->{CommandText} = "<$sBase>;$sFilter;$sAttributes;$sScope";

my $ADOrecordset = $ADOCommand->Execute();


while(!$ADOrecordset->{EOF}){
    my $dn = $ADOrecordset->Fields("distinguishedName")->{Value};
    printf("distinguishedName => %-40s\n", $dn);
    $ADOrecordset->MoveNext();
}

$ADOConnection->Close();
$ADOrecordset->Close();
```
