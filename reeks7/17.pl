# An integer that indicates that the attribute is a linked attribute. An even integer is a forward link and an odd integer is a back link.
# Gevonden in characteristics of attributes

# We willen de bijhorende forward en backward linked attributes
# We gaan dat zo fixen
# Eerst alle forward gaan ophalen: Dat zijn die waar het eerste bit op een 0 staat
# Daarna de backward gaan opahelen

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

# We moeten de linkID hebben waarbij het minst grootte bit op 0 staat
# Dit kunnen we doen met een nor porrt
# 0 - 0 = 0 ==> !0 = 1
my $sFilter = "(&(objectclass=attributeSchema)(linkID=*)(!(linkID:1.2.840.113556.1.4.804:=1)))";
my $sAttributes = "linkID,adsPath,name"; 
my $sScope      = "subtree";

$ADOcommand->{CommandText} = "<$sBase>;$sFilter;$sAttributes;$sScope";

my $ADOrecordset = $ADOcommand->Execute();
Win32::OLE->LastError()&& die (Win32::OLE->LastError());

while(!$ADOrecordset->{EOF}){
    my $ForwardLinkName = $ADOrecordset->Fields("Name")->{Value};
    my $forwardID = $ADOrecordset->Fields("linkID")->{Value};
    my $backID= $forwardlinkID + 1;


    my $sFilter = "(&(objectclass=attributeSchema)(linkID=*)(linkID=$forwardID))";
    $ADOcommand->{CommandText} = "<$sBase>;$sFilter;$sAttributes;$sScope";

    my $ADOrecordset2 = $ADOcommand->Execute();
    Win32::OLE->LastError()&& die (Win32::OLE->LastError());

    $backLinkName = $ADOrecordset2->Fields("name")->{Value};
    $LinkAttributes{$ForwardLinkName} = $backLinkName;


    $ADOrecordset->MoveNext();
}

while( ($key, $value) = each(%LinkAttributes) ){
    print ("forward: $key => backward: $value\n");
}