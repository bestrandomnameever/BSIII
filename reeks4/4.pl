# Die directory is Win32_Directory.Name="c:\\". Kan je vinden door eerst te zoeken op _LogicalDisk
use Win32::OLE;
Win32::OLE->Option(Warn => 3);

$ComputerName = ".";
$NameSpace = "root/cimv2";

$ClassName = 'Win32_Directory.Name="c:\\\\"';

$moniker = "winmgmts://$ComputerName/$NameSpace:$ClassName";
$Directory = Win32::OLE->GetObject($moniker);


printf("FileType => %s\n", $Directory->FileType);