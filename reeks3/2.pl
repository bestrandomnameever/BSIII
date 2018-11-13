# Overgeërfde attributen hebben een gele pijl links van de attribuutnaam
# Attributen van de klasse zelf hebben zo'n wit icoontje ernaast

# Overerving kan je zien in de DYNASTY-ARRAY of in de boomstructuur links 

# Icoontje van het sleutelattribuut is een sleutel: Hier is dat DeviceID

# DeviceID is al opgenomen in CIM_LogicalDevice. Dit vind je door gewoon omhoog te gaan in de klassen en te kijken of DeviceID
# er al bijstaat tot er geen gele pijl meer staat. Dan is dat attribuut immers niet meer overgeërfd.

# Hoe tel je ze? Met propertycount weet je dat er 59 zijn
# De superklasse heet CIM_PCVideoController (dit kan je weten door de _DYNASTY-array te bekijken)
# 59 - 41 = 18