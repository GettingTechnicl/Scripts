Import-Csv -Path "C:\temp\AzureADUsersPwd.csv" | foreach {

$users = get-msoluser | select userprincipalname,objectid | where UserPrincipalName -like $_.UserPrincipalName 

Set-MsolUserPassword -UserPrincipalName $_.UserPrincipalName -NewPassword $_.Password -ForceChangePassword $False }