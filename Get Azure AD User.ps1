Get-AzureADUser -All $true | 
  Select-Object *,@{label="Manager";expression={(Get-AzureADUserManager -ObjectId $_.ObjectID).displayname}} | 
  Export-Csv c:\temp\file.csv
Get-AzureADUser -All $true | 
  Select-Object *,@{label="Manager";expression={(Get-AzureADUserManager -ObjectId $_.ObjectID).displayname}},  @{n = 'assignedlicenses1'; e = { ($_.assignedlicenses.SkuId) -join "," } },
    @{n = 'assignedplans1'   ; e = { ($_.assignedplans.service) -join "," } }|
  Export-Csv c:\temp\file.csv -noTypeInformation

