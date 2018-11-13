# Alle bestanden van de huidige map oplijsten 
dir -Path "." -File;

# Toon het aantal directories/files per soort
dir -Path "." | Group-Object -Property Mode 

# Na 01/10/2017
$Date = Get-Date -Date "01/10/2017" -Format "dd/MM/yyyy"
dir -Path "." | where {$_.LastWriteTime -ge $Date} | select Name, LastWriteTime