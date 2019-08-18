use Win32::OLE qw(in);
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Data::Dumper;

my $Service = Win32::OLE->GetObject('winmgmts://./root/cimv2');
our $StdRegProv = $Service->Get('StdRegProv');


my %RootKey = ( HKEY_CLASSES_ROOT   => 0x80000000
              , HKEY_CURRENT_USER   => 0x80000001
              , HKEY_LOCAL_MACHINE  => 0x80000002
              , HKEY_USERS          => 0x80000003
              , HKEY_CURRENT_CONFIG => 0x80000005
              , HKEY_DYN_DATA       => 0x80000006 );


# We gaan EnumKey gebruiken van het register

our $InParameters = $StdRegProv->Methods_->{EnumKey}{InParameters};
our $InParameters->{hDefKey} = $RootKey{HKEY_LOCAL_MACHINE};

sub PrintTak {
    (my $Key, my $Depth) = @_;
    $InParameters->{sSubKeyName} = $Key;
    my $OutParams = $StdRegProv->ExecMethod_(EnumKey, $InParameters);
    return if $OutParams->{ReturnValue}; # Als niet 0 is dan gaat er iets fout
    return if !defined $OutParams->{sNames};
    foreach $subKey (@{$OutParams->{sNames}}){
        print " " x ($Depth * 5), "$subKey\n";
        PrintTak($Key . "\\" . $subKey, $Depth+1);
    }
}

sub PrintSubKeys {
    ( my $Key ) = @_;
    $InParameters->{sSubKeyName} = $Key;
    print "$Key\n";
    PrintTak($Key, 1);
}
PrintSubKeys('SYSTEM\\Currentcontrolset\\Services\\tcpip');







