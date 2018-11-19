#Get-Help Group-Object -parameter *
Get-Command -CommandType cmdlet | Group-Object -Property verb | sort -Property Count -Descending | select -first 10;