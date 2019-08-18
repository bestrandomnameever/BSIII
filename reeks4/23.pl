use Win32::OLE qw(in);
use Data::Dumper;

$fso = Win32::OLE->new('Scripting.FileSystemObject');
@ARGV = ('.');
sub DirectorySize{
    my $Size = 0;
    (my $directory, my $depth) = @_;
    print $directory->{Name} . "\n";
    if($depth > 0){
        $depth--;
        foreach(in $directory->Associators_(undef, 'Win32_Directory', 'PartComponent')){
            $Size += DirectorySize($_, $depth);
        }
    }

    foreach(in $directory->Associators_('CIM_DirectoryContainsFile')){
        printf("    %s => %s\n", $_->{Description}, $_->{FileSize});
        $Size+= $_->{FileSize};
    }
    return $Size;
}   

(my $path = lc($fso->GetAbsolutePathName(shift))) =~ s/\\/\\\\/g;

my $Service = Win32::OLE->GetObject('winmgmts://./root/cimv2');
( my $Directory ) = in $Service->ExecQuery("select * from Win32_Directory where Name = '$path'");

my $Totaal = DirectorySize($Directory, 2);
print Dumper $Totaal;



