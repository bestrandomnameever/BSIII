use Win32::OLE::Const "^Microsoft Excel";
use Win32::OLE::Const ".*CDO";

# Nu de libraries ingeladen zijn kan je constanten gaan zoeken
# In OLEVIEW zoek je naar Microsoft Excel 12.0 Object Library
# Dubbelklik en in de lijst links vind je "typedef enum Constants". 
# Hier staan constants in

printf("xlNextToAxis: %s\n", xlNextToAxis);
