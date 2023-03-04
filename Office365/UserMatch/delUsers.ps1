# 365 Authentication
$credential = Get-Credential
Connect-MsolService -Credential $credential

# Import CSV
$Datas = import-csv C:\temp\users.csv
	 foreach($Data in $Datas){
	 	Remove-MSOLuser -UserPrincipalName "$Data.UserPrincipalName" -RemoveFromRecycleBin
	}
	
	
	## AI GENERATED BELOW ##
	
# Connect to Office 365 using the latest MSOnline V2 module
Connect-MsolService

# Import CSV and remove users
Import-Csv -Path "C:\temp\users.csv" | ForEach-Object {
    Write-Host "Removing user $($_.UserPrincipalName)"
    $user = Get-MsolUser -UserPrincipalName $_.UserPrincipalName
    Remove-MsolUser -ObjectId $user.ObjectId -RemoveFromRecycleBin -Force -Confirm:$false
}

