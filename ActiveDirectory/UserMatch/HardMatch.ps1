# Hard Match
# https://www.itpromentor.com/soft-vs-hard-match/

# README
# Create .csv with 2 columns -> username | UserPrincipalName
# update import-csv to point to correct path for created .csv file
# run this script on the AD

# 365 Authentication
$credential = Get-Credential
Connect-MsolService -Credential $credential

# Import CSV
$Datas = import-csv C:\temp\users.csv
# Correct immutable ID's for users    
 foreach($Data in $Datas){
#Param(
#$username
#)
$365User="$data.UserPrincipalName"
$guid=(get-ADUser $data.username).Objectguid
$immutableID=[system.convert]::ToBase64String($guid.tobytearray())
Set-MsolUser -UserPrincipalName "$365User" -ImmutableId $immutableID 
 }