use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Data::Dumper;
$DateTime = Win32::OLE->new("WbemScripting.SWbemDateTime");

@properties = ( );

sub newline {
    ( my $count ) = @_;
    $count = 1 if !defined $count;
    print "\n" x $count;
}

sub checkQualifier {
    ( my $Class ) = @_;
    $Qualifiers = $Class->{Qualifiers_};
    for $Qualifier (in $Qualifiers){
        $hash{lc $Qualifier->{Name}}++ if $Qualifier->{Name} =~ /abstract|association|dynamic|singleton/i;
    }
}

my $ComputerName  = ".";
my $NameSpace = "root/cimv2";
my $WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");

for (in $WbemServices->SubClassesOf(undef)){
    checkQualifier($_);
}

while ( ($key, $value) = each(%hash) ){
    printf("%s : %d\n", $key, $value);
}