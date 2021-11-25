Get-AzureADUser -All $true | 
  Select-Object *,@{label="Manager";expression={(Get-AzureADUserManager -ObjectId $_.ObjectID).displayname}} | 
  Export-Csv c:\temp\file.csv
