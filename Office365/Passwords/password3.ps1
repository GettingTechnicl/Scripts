Import-Csv -Path "C:\temp\AzureADUsersPwd.csv" | foreach {
$userUPN=$_.'UserPrincipalName'
$newPassword=$_.'Password'
$secPassword = ConvertTo-SecureString $newPassword -AsPlainText -Force
Set-AzureADUserPassword -ObjectId  $userUPN -Password $secPassword}