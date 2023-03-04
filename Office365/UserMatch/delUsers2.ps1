# 365 Authentication
$credential = Get-Credential
Connect-MsolService -Credential $credential

# Import CSV
Import-Csv C:\temp\users.csv | ForEach($Data in $Datas){Remove-MSOLuser -UserPrincipalName "$Data.UserPrincipalName" -RemoveFromRecycleBin}