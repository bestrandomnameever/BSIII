use Win32::OLE 'in';
use Win32::OLE::Const;
use Data::Dumper;

%constanten = %{Win32::OLE::Const->Load('Active DS Type Library')};

while( (my $key, my $value) = each(%constanten) ){
    print "$key ==> $value\n" if $key =~ /^ADS_/;
}