#PROGID: {76A64158-CB41-11D1-8B02-00600806D9B6}
# Je komt uit op WBEM Scripting Locator
# Hierin staat een TypeLib met id {565783C6-CB41-11D1-8B02-00600806D9B6}
# TypeLibrary: Microsoft WMI Scripting V1.2 Library

use Win32::OLE::Const;

#Voor warnings
Win32::OLE->Option(Warn => 3);

$object = Win32::OLE->new('WbemScripting.SWbemLocator.1');
%constanten = %{Win32::OLE::Const->Load($object)};

while( ($key, $value) = each %constanten ){
    printf("%s => %s\n", $key, $value);
}

# Of 

use Win32::OLE::Const "Microsoft WMI Scripting V1.2 Library";
printf("wbemFlagUseCurrentTime: %s\n", wbemFlagUseCurrentTime);