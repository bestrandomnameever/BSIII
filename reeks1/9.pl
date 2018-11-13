# Inladen van const
use Win32::OLE::Const;

#gedrag simuleren van argument meegeven
@ARGV = ("Microsoft Excel 12.0 Object Library");
$librarynaam = $ARGV[0];

# In de hash worden nu alle constanten gestoken van de library
my $hash = Win32::OLE::Const->Load($librarynaam);

while( ($key, $value) = each %{$hash} ){
    print "$key => $value\n";
}