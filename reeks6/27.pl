use bind_object;
use Data::Dumper;
use Win32::OLE::Variant;

my %E_ADS = (
    BAD_PATHNAME            => Win32::OLE::HRESULT(0x80005000),
    UNKNOWN_OBJECT          => Win32::OLE::HRESULT(0x80005004),
    PROPERTY_NOT_SET        => Win32::OLE::HRESULT(0x80005005),
    PROPERTY_INVALID        => Win32::OLE::HRESULT(0x80005007),
    BAD_PARAMETER           => Win32::OLE::HRESULT(0x80005008),
    OBJECT_UNBOUND          => Win32::OLE::HRESULT(0x80005009),
    PROPERTY_MODIFIED       => Win32::OLE::HRESULT(0x8000500B),
    OBJECT_EXISTS           => Win32::OLE::HRESULT(0x8000500E),
    SCHEMA_VIOLATION        => Win32::OLE::HRESULT(0x8000500F),
    COLUMN_NOT_SET          => Win32::OLE::HRESULT(0x80005010),
    ERRORSOCCURRED          => Win32::OLE::HRESULT(0x00005011),
    NOMORE_ROWS             => Win32::OLE::HRESULT(0x00005012),
    NOMORE_COLUMNS          => Win32::OLE::HRESULT(0x00005013),
    INVALID_FILTER          => Win32::OLE::HRESULT(0x80005014),
    INVALID_DOMAIN_OBJECT   => Win32::OLE::HRESULT(0x80005001),
    INVALID_USER_OBJECT     => Win32::OLE::HRESULT(0x80005002),
    INVALID_COMPUTER_OBJECT => Win32::OLE::HRESULT(0x80005003),
    PROPERTY_NOT_SUPPORTED  => Win32::OLE::HRESULT(0x80005006),
    PROPERTY_NOT_MODIFIED   => Win32::OLE::HRESULT(0x8000500A),
    CANT_CONVERT_DATATYPE   => Win32::OLE::HRESULT(0x8000500C),
    PROPERTY_NOT_FOUND      => Win32::OLE::HRESULT(0x8000500D) );

# Lijst van LDAP attributen
my @attributen=qw(accountNameHistory aCSPolicyName adminCount adminDescription adminDisplayName
    allowedAttributes allowedAttributesEffective allowedChildClasses allowedChildClassesEffective
    altSecurityIdentities assistant badPwdCount bridgeheadServerListBL c canonicalName cn co
    codePage comment company countryCode createTimeStamp defaultClassStore department description
    desktopProfile destinationIndicator directReports displayName displayNamePrintable distinguishedName
    division dSCorePropagationData dynamicLDAPServer employeeID extensionName facsimileTelephoneNumber
    flags fromEntry frsComputerReferenceBL fRSMemberReferenceBL fSMORoleOwner garbageCollPeriod
    generationQualifier givenName groupPriority groupsToIgnore homeDirectory homeDrive homePhone
    homePostalAddress info initials instanceType internationalISDNNumber ipPhone isCriticalSystemObject
    isDeleted isPrivilegeHolder l lastKnownParent legacyExchangeDN localeID logonCount mail
    managedObjects manager masteredBy memberOf mhsORAddress middleName mobile modifyTimeStamp
    mS-DS-ConsistencyChildCount msNPAllowDialin msNPCallingStationID msNPSavedCallingStationID
    msRADIUSCallbackNumber msRADIUSFramedIPAddress msRADIUSFramedRoute msRADIUSServiceType
    msRASSavedCallbackNumber msRASSavedFramedIPAddress msRASSavedFramedRoute name netbootSCPBL
    networkAddress nonSecurityMemberBL o objectCategory objectClass objectVersion operatorCount
    otherFacsimileTelephoneNumber otherHomePhone otherIpPhone otherLoginWorkstations otherMailbox
    otherMobile otherPager otherTelephone otherWellKnownObjects ou pager personalTitle
    physicalDeliveryOfficeName postalAddress postalCode postOfficeBox preferredDeliveryMethod preferredOU
    primaryGroupID primaryInternationalISDNNumber primaryTelexNumber profilePath proxiedObjectName
    proxyAddresses queryPolicyBL revision rid sAMAccountName sAMAccountType scriptPath
    sDRightsEffective seeAlso serverReferenceBL servicePrincipalName showInAddressBook
    showInAdvancedViewOnly siteObjectBL sn st street streetAddress subRefs subSchemaSubEntry
    systemFlags telephoneNumber textEncodedORAddress title url userAccountControl userParameters
    userPrincipalName userSharedFolder userSharedFolderOther userWorkstations USNIntersite wbemPath
    wellKnownObjects whenChanged whenCreated wWWHomePage x121Address);

sub printAttribute {
    (my $Attribute, my $Object ) = @_;
    my $arrayref = $Object->GetEx("$Attribute");

    if(Win32::OLE->LastError() == $E_ADS{PROPERTY_NOT_FOUND}){
        # Opvragen met GetInfoEx
        $Object->GetInfoEx("$Attribute");
        $arrayref = $Object->GetEx("$Attribute");
    } 

    if(Win32::OLE->LastError() == $E_ADS{PROPERTY_NOT_FOUND}){
        print "$Attribute is niet ingevuld\n";
    } else {
        print "Waarden voor $Attribute :\n";
        print "$_\n" foreach @{$arrayref};
    }

    print "\n\n";
}
my $Object = bind_object::bind_object("CN=Administrator,CN=Users,DC=iii,DC=hogent,DC=be");
printAttribute($_, $Object) foreach @attributen;
