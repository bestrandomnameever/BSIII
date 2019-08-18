use Win32::OLE::Const;
use Data::Dumper;

%constanten = %{Win32::OLE::Const->Load("Active DS Type Library")};

@ARGV = ("_UF_");

print "$_ ==> $constanten{$_}\n" foreach (grep { /$ARGV[0]/ } keys %constanten); 