#Read user details from the CSV file
$CSVRecords = Import-CSV "C:\Temp\AzureADUsersPwd.csv"
$i = 0;
$TotalRecords = $CSVRecords.Count
  
#Array to add the status result
$UpdateResult=@()
  
#Iterate users one by one and set the password 
Foreach($CSVRecord in $CSVRecords)
{
#Get Object ID for each user
$upN = '$CSVRecord.UserPrincipalName'
$UserId = Get-AzureADUser -filter "userPrincipalName eq '$upN'" | ft ObjectID

#Convert the password to a secure string 
$NewPassword = ConvertTo-SecureString '$CSVRecord.Password' -AsPlainText -Force
 
  
$i++;
Write-Progress -activity "Processing $UserId " -status "$i out of $TotalRecords users completed"
  
try
{
#Set the password value and set force change password at next login flag
Set-AzureADUserPassword –ObjectId "$UserId" –Password $NewPassword -ForceChangePasswordNextLogin $False
$ResetStatus = "Success"
}
catch
{
$ResetStatus = "Failed: $_"
}
  
#Add reset password status
$UpdateResult += New-Object PSObject -property $([ordered]@{
User = $UserId
ResetPasswordStatus = $ResetStatus
})
}
  
#Display the reset password status result
$UpdateResult | Select User,ResetPasswordStatus | FT
  
#Export the reset password status report to a CSV file
$UpdateResult | Export-CSV "C:\temp\ResetPasswordStatus.CSV" -NoTypeInformation -Encoding UTF8
