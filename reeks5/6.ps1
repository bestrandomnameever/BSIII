#Initialiseer een variabele met een tijd
$tijd=(Get-ChildItem | select -first 1).LastWriteTime;
#zoek alle properties en methodes
$tijd | Get-Member | foreach{$_.MemberType.toString() + "`t`t`t==> " + $_.Name};

# Returntype is DayOfWeek, deze moet je dus toString() doen zodat je hem kan vergelijken 
dir -Path "." -file | where{$_.LastWriteTime.DayOfWeek.toString() -eq "Sunday"} |
 foreach { 
    Write-Host $_.Name, $_.LastWriteTime.TimeOfDay
}
 
# Zonder foreach:

dir -Path "." -file | where {$_.LastWriteTime.DayOfWeek -eq "Sunday"} | 
        select  Name,@{Name="Last Access";Expression={$_.LastWriteTime.TimeOfDay}}
