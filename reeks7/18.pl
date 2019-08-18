use bind_object;
use Data::Dumper;
use Win32::OLE::Variant;

@ARGV = ("user");

foreach(@ARGV){
    
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


    my $domein = bind_object::bind_object( $RootDSE->Get("defaultNamingContext") );
    my $sBase = $domein->{adspath};

    my $sFilter = "(objectCategory=$_)";
    my $sAttributes = "objectClass,objectCategory"; 
    my $sScope      = "subtree";

    $ADOcommand->{CommandText} = "<$sBase>;$sFilter;$sAttributes;$sScope";

    my $ADOrecordset = $ADOcommand->Execute();
    Win32::OLE->LastError()&& die (Win32::OLE->LastError());

    while(!$ADOrecordset->{EOF}){
        printf("%-20s : %-20s\n", $_, $ADOrecordset->{$_}->{Value} =~ /ARRAY/i ? join ", ", @{$ADOrecordset->{$_}->{Value}} : $ADOrecordset->{$_}->{Value} ) foreach split(",",$sAttributes);
        print "\n" x 3;
        $ADOrecordset->MoveNext();
    }
}