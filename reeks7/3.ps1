$ADOConnection = New-Object -ComObject "ADODB.Connection";
$ADOconnection.Provider = "ADsDSOObject";

$ADOconnection.Properties.Item("User ID").Value="Cedric Vanhaverbeke";  # vul in of zet in commentaar op school
$ADOconnection.Properties.Item("PassWord").Value="Cedric Vanhaverbeke"; # vul in of zet in commentaar op school
$ADOconnection.Properties.Item("Encrypt Password").Value=1;

$ADOconnection.Open();                                   # mag je niet vergeten

$ADOcommand = new-Object -ComObject "ADODB.Command"
$ADOcommand.ActiveConnection      = $ADOconnection;         
$ADOcommand.Properties.Item("Page Size").value = 20;

