# Met get-member kan je alle members opvragen van iets. Best handig
Get-Command | select -first 1 | Get-Member -MemberType Property;

Get-Command | where {$_.Verb -match "^$"} | select Name;
