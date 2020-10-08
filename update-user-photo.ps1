$UserCredential = Get-Credential
$DomainName
#Update the user photos path here. Name of the file should be username of the Office365 user.
$path= 'C:\Staff\'
# Upload the photo
#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/?proxyMethod=RPS -Credential $UserCredential -Authentication Basic -AllowRedirection
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection

Import-PSSession $Session


$Images = Get-ChildItem $path
$Images |Foreach-Object{
$Identity = ($_.Name.Tostring() -split "\.")[0]+$DomainName
$PictureData = $path+$_.name
Set-UserPhoto -Identity $Identity -PictureData ([System.IO.File]::ReadAllBytes($PictureData)) -Confirm:$false }
                        

                        
