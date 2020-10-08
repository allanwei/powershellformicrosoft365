
Connect-AzureAD
$Result = @()

$AllUsers= Get-AzureADUser -All $true |Where {$_.UserType -eq 'Member' -and $_.AssignedLicenses -ne $null} | Select-Object -Property Displayname,UserPrincipalName,JobTitle,PhysicalDeliveryOfficeName,Department

$TotalUsers = $AllUsers.Count

$i = 1 

$AllUsers | ForEach-Object {
$User = $_
Write-Progress -Activity "Processing $($_.Displayname)" -Status "$i out of $TotalUsers completed"
$Licenses=Get-AzureADUser -ObjectId $User.UserPrincipalName | Select -ExpandProperty AssignedLicenses
$managerObj = Get-AzureADUserManager -ObjectId $User.UserPrincipalName
$Result += New-Object PSObject -property @{ 
UserName = $User.DisplayName
UserPrincipalName = $User.UserPrincipalName
ManagerName = if ($managerObj -ne $null) { $managerObj.DisplayName } else { $null }
ManagerMail = if ($managerObj -ne $null) { $managerObj.Mail } else { $null }
Office = $User.PhysicalDeliveryOfficeName
JobTitle=$User.JobTitle
Department=$User.Department
}
$i++
}
$Result | Select UserName,UserPrincipalName,JobTitle,Department,Office,ManagerName,ManagerMail  |
Export-CSV "C:\O365UsersManagerInfo.CSV" -NoTypeInformation -Encoding UTF8
