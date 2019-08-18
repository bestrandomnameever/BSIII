use bind_object;
use Data::Dumper;
use Win32::OLE::Variant;

my $ADOconnection = Win32::OLE->new("ADODB.Connection");
$ADOconnection->{Provider} = "ADsDSOObject";

if ( uc($ENV{USERDOMAIN}) ne "III") { 
$ADOconnection->{Properties}->{"User ID"} = "Interim F";
$ADOconnection->{Properties}->{"Password"} = "Interim F";
$ADOconnection->{Properties}->{"Encrypt Password"} = True;
}

$ADOconnection->Open();      

my $ADOcommand = Win32::OLE->new("ADODB.Command");
$ADOcommand->{ActiveConnection} = $ADOconnection;   
$ADOcommand->{Properties}->{"Page Size"} = 20;

my $RootDSE=bind_object::bind_object("RootDSE");
$RootDSE->getinfo();


my $domein = bind_object::bind_object( $RootDSE->Get("schemaNamingContext") );
my $sBase = $domein->{adspath};

my $sFilter = "(&(objectCategory=classSchema)(defaultObjectCategory=*))";
my $sAttributes = "defaultObjectCategory,ldapDisplayName,objectCategory"; 
my $sScope      = "OneLevel";

$ADOcommand->{CommandText} = "<$sBase>;$sFilter;$sAttributes;$sScope";

my $ADOrecordset = $ADOcommand->Execute();
Win32::OLE->LastError()&& die (Win32::OLE->LastError());

while(!$ADOrecordset->{EOF}){
    $ADOrecordset->Fields("defaultObjectCategory")->{Value} =~ /.*?=(.*?),/;
    $defaultObjectCategory = $1;
    $ldapDisplayName = $ADOrecordset->Fields("ldapDisplayName")->{Value};
    $objectCategory = $ADOrecordset->Fields("objectCategory")->{Value};



    my $classes = {};
    $classes->{objectCategory} = $objectCategory;
    $classes->{ldapDisplayName} = $ldapDisplayName;

    push @{$hash->{$defaultObjectCategory}}, $classes;

    $ADOrecordset->MoveNext();
}

while( ($key, $value) = each (%{$hash})){
    if (scalar @{$value} gt 1){
        print  "$key\nmet volgende waarden\n";
        print Dumper $value;
    }
}