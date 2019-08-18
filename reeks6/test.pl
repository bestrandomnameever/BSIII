use Win32::OLE::Const 'Active DS Type Library';
use Win32::OLE 'in';
use Win32::OLE::Variant;
use Data::Dumper;
use bind_object;

my %Errors = (
    E_ADS_PROPERTY_NOT_FOUND => Win32::OLE::HRESULT(0x8000500D)
);

sub printProperty {

    my $PropertyName = shift;
    $ArrayOfValues = $Object->GetEx($PropertyName);

    if(Win32::OLE->LastError() eq $Errors{E_ADS_PROPERTY_NOT_FOUND}){
        printLijn($PropertyName, "<Niet ingesteld>");
        next;
    }

    my $abstracteProperty = bind_object::bind_object("schema/$PropertyName");

    foreach (@{$ArrayOfValues}){
            my $value;
            #if ($_ eq "OctetString") {
            #    $value = sprintf ("%*v02X ","", $_);
            #} else 
            if ($_ eq "ObjectSecurityDescriptor") {
                $value=$_->{owner};
            } else {
                $value = $_;
            }

            printLijn($PropertyName,$value);

        }
    }


sub printLijn {
    (my $PropertyName, my $Value) = @_;
    $AantalVoorkomens{$PropertyName}++;
    printf("%-55s%s\n", $AantalVoorkomens{$PropertyName} == 1 ? $PropertyName : "", $Value);
}

$RootDSE = bind_object::bind_object("RootDSE");
die "Object niet gevonden" if Win32::OLE->LastError();

$Object = bind_object::bind_object("OU=iii," . $RootDSE->Get(defaultNamingContext));
die "Object niet gevonden" if Win32::OLE->LastError();

$SchemaObject = bind_object::bind_object($Object->{Schema});
die "Object niet gevonden" if Win32::OLE->LastError();

@Mandatory = @{$SchemaObject->{MandatoryProperties}};

# Ophalen van alle mandatory properties zodat we dat later niet meer moeten doen
$Object->GetInfoEx(\@Mandatory, 0);

@Optional = @{$SchemaObject->{OptionalProperties}};
$Object->GetInfoEx(\@Optional, 0);

printf("MANDATORY PROPERTIES\n%s\n", "-" x 50);
foreach my $property (@Mandatory){
    printProperty($property);
}

print "\n";

printf("OPTIONELE PROPERTIES\n%s\n", "-" x 50);
foreach my $property (@Optional){
    printProperty($property);
}





