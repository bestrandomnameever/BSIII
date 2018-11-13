use Win32::OLE 'in';
use Win32::OLE::Const 'Microsoft WMI Scripting ';
use Data::Dumper;

my $ComputerName  = ".";
my $NameSpace = "root/cimv2";
my $WbemServices = Win32::OLE->GetObject("winmgmts://$ComputerName/$NameSpace");

# Die wbemFlagsUseAmendedQualifiers is superbelangrijk! Anders krijg je geen Values terug en kan je geen hash opbouwen
my $NetworkAdapterSchema = $WbemServices->Get("Win32_NetworkAdapter", wbemFlagUseAmendedQualifiers);

sub getHash {
    (my $Schema, my $prop) = @_;
    %hash;
    @hash{ @{$Schema->Properties_->Item($prop)->Qualifiers_->{"ValueMap"}->{Value}} } = @{$Schema->Properties_->Item($prop)->Qualifiers_->{"Values"}->{Value}};
    return \%hash;
}

my %AvailabilityHash = %{getHash($NetworkAdapterSchema, "Availability")};
my %NetConnectionStatusHash = %{getHash($NetworkAdapterSchema, "NetConnectionStatus")};

my $Adapters = $WbemServices->ExecQuery("SELECT * from Win32_NetworkAdapter where NetConnectionStatus is not null");

for my $Adapter (in $Adapters){
    
    printf("Name: %s\nType: %s\nToestand:%s\nBeschikbaarheid: %s\nStatus: %s\nMAC address: %s\nAdapter Service name: %s\nLast Reset: %s",
        $Adapter->Properties_->Item("Name")->{Value},
        $Adapter->Properties_->Item("AdapterType")->{Value},
        $AvailabilityHash{$Adapter->Properties_->Item("Availability")->{Value}},
        $AvailabilityHash{$Adapter->Properties_->Item("NetConnectionStatus")->{Value}},
        # Zo kan je ook properties ophalen
        $Adapter->{MACAddress},
        $Adapter->{ServiceName}, 
        $Adapter->{TimeOfLastReset});

    #Recource Informatie
    $Query="ASSOCIATORS OF {Win32_NetworkAdapter='$Adapter->{Index}'} 
               WHERE AssocClass=Win32_AllocatedResource";
    my $AdapterResourceInstances = $WbemServices->ExecQuery ($Query);
    foreach $AdaptResInstance (in $AdapterResourceInstances) {
        my $className=$AdaptResInstance->{Path_}->{Class};
        printf "%s: %s\n", "IRQ resource", $AdaptResInstance->{IRQNumber} if $className eq "Win32_IRQResource";
        printf "%s: %s\n", "DMA channel", $AdaptResInstance->{DMAChannel} if $className eq "Win32_DMAChannel";
        printf "%s: %s\n", "I/O Port", $AdaptResInstance->{Caption}       if $className eq "Win32_PortResource";
        printf "%s: %s\n", "Memory address", $AdaptResInstance->{Caption} if $className eq "Win32_DeviceMemoryAddress";
    }

    my $AdapterInstance = $WbemServices->Get ("Win32_NetworkAdapterConfiguration=$AdapterInstance->{Index}");
    next unless $AdapterInstance->{IPEnabled};

    if ($AdapterInstance->{DHCPEnabled}) {
       printf "%s: %s\n", "DHCP expires", $AdapterInstance->{DHCPLeaseExpires};
       printf "%s: %s\n", "DHCP obtained", $AdapterInstance->{DHCPLeaseObtained};
       printf "%s: %s\n", "DHCP server", $AdapterInstance->{DHCPServer};
    }

    # Dit zijn arrays
    printf "%s: %s\n", "IP address(es)", (join ",",@{$AdapterInstance->{IPAddress}});
    printf "%s: %s\n", "IP mask(s)", (join ",",@{$AdapterInstance->{IPSubnet}});
    printf "%s: %s\n", "IP connection metric", $AdapterInstance->{IPConnectionMetric};
    printf "%s: %s\n", "Default Gateway(s)",(join ",",@{$AdapterInstance->{DefaultIPGateway}});
    printf "%s: %s\n", "Dead gateway detection enabled", $AdapterInstance->{DeadGWDetectEnabled};

    printf "%s: %s\n", "DNS registration enabled", $AdapterInstance->{DomainDNSRegistrationEnabled};
    printf "%s: %s\n", "DNS FULL registration enabled", $AdapterInstance->{FullDNSRegistrationEnabled};
    printf "%s: %s\n", "DNS search order", (join ",",@{$AdapterInstance->{DNSServerSearchOrder}});
    printf "%s: %s\n", "DNS domain", $AdapterInstance->{DNSDomain};
    printf "%s: %s\n", "DNS domain suffix search order", $AdapterInstance->{DNSDomainSuffixSearchOrder};
    printf "%s: %s\n", "DNS enabled for WINS resolution", $AdapterInstance->{DNSEnabledForWINSResolution};
}


