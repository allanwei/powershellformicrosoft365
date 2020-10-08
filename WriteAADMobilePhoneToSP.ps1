
param(
$UserId,
$PWD
)
Try {
if(Get-Module -ListAvailable -Name MSOnline){
    Write-Output "Import-Module MSOnline ..."
    Import-Module MSOnline
}
else{
    Write-Output "Import-Module MSOnline ..."
    Install-Module MSOnline -Scope CurrentUser -Force
    Import-Module MSOnline
}
if (Get-Module -ListAvailable -Name SharePointPnPPowerShellOnline) {
    Write-Output "Import-Module SharePointPnPPowerShellOnline ..."
    #Import-Module SharePointPnPPowerShellOnline
} 
else {
     Write-Output "Import-Module SharePointPnPPowerShellOnline ..."
    Install-Module SharePointPnPPowerShellOnline -Scope CurrentUser -Force
    #Import-Module SharePointPnPPowerShellOnline
}



$spoAdminUrl = "https://xxxxxxx-admin.sharepoint.com"  
$overwriteExistingSPOUPAValue = "True"  #Always overwrite
$pass =ConvertTo-SecureString $PWD -AsPlainText -Force
$mycred=New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserId,$pass
$O365Cred = Get-Credential $mycred

# Get credentials of account that is AzureAD Admin and SharePoint Online Admin
#$credential = Get-AutomationPSCredential -Name "SharepointAdmin";

Write-Output "START process..."


    # Connect to AzureAD    
    Connect-MsolService -Credential $O365Cred
    
    # Connect to SharePointOnline
    Connect-PnPOnline -Url $spoAdminUrl -Credentials $O365Cred #$credential  

    # Get all Licensed AzureAD Users
    $AzureADUsers = Get-MSolUser -All|?{($_.UserType -eq "Member") -and ($_.IsLicensed)}
    
    ForEach ($AzureADUser in $AzureADUsers) {
     
         

        $adUserMobilePhone = $AzureADUser.MobilePhone
        $targetUPN = $AzureADUser.UserPrincipalName.ToString()
        
		
		
        # Check to see if the AzureAD User has a MobilePhone specified, if nothing, set it blank to overwrite
        if ([string]::IsNullOrEmpty($adUserMobilePhone)) {
         $adUserMobilePhone = "";
        }
        
        $adUserOfficeNumber = $AzureADUser.Office
        #Write-Output "AD User office $adUserOfficeNumber for $targetUPN"
		# Check if user is working in office or field
        $workLoc = "";
        if ([string]::IsNullOrEmpty($adUserOfficeNumber)) {
            $workLoc = "";  #catch case adUserOfficeNumber is null, stop location check here
        } elseif ($adUserOfficeNumber.StartsWith("field","CurrentCultureIgnoreCase")) {
            $workLoc = "Field";
        } elseif ($adUserOfficeNumber.StartsWith("office","CurrentCultureIgnoreCase") ) {
            $workLoc = "Office";
        }
        #Write-Output "AD User Loc $workLoc for $targetUPN"
        
        $targetSPOUserAccount = ("i:0#.f|membership|" + $targetUPN)
        $targetSPOUserProfile = Get-PnPUserProfileProperty -Account $targetSPOUserAccount -ErrorAction SilentlyContinue
        
        if ($null -ne $targetSPOUserProfile) {
            # Get current spo user mobile # for verification
            $spoUserCellPhone = $targetSPOUserProfile.UserProfileProperties.CellPhone
            #Write-Output "SPO User existing cell phone $spoUserCellPhone"

            # If target property is empty let's populate it
            if ($adUserMobilePhone -ne $spoUserCellPhone) {
                if ($overwriteExistingSPOUPAValue -eq "True") {                
                    Write-Output "Assign SPS user cellphone $spoUserCellPhone with new number $adUserMobilePhone for $targetUPN" -Debug
                    Set-PnPUserProfileProperty -Account $targetSPOUserAccount -PropertyName "CellPhone" -Value $adUserMobilePhone                
                } else {
                    # Not going to overwrite existing property value
                    Write-Output "SPS Property is set to NOT overwrite. We're preserve existing cellphone $spoUserCellPhone (New cellphone $adUserMobilePhone not write)"  
                }
            } 

            $spoUserLocation = $targetSPOUserProfile.UserProfileProperties["SPS-Location"]
            #Write-Output "SPO User existing location $spoUserLocation"
            if ($workLoc -ne $spoUserLocation) {
                 if ($overwriteExistingSPOUPAValue -eq "True") {                
                    Write-Output "Assign SPS user location $spoUserLocation with new location $workLoc for $targetUPN"  -Debug
                    Set-PnPUserProfileProperty -Account $targetSPOUserAccount -PropertyName "SPS-Location" -Value $workLoc                
                } else {
                    # Not going to overwrite existing property value
                    Write-Output "SPS Property is set to NOT overwrite. We're preserve existing location $spoUserLocation (New location $workLoc not write)"  
                }
            }


        } else {
            # SPO User is empty, nothing to do here
            Write-Output "Target SPO User $targetUPN is not found." 
        }
    }
       
    
    Write-Output "END Process!"
    
}
Catch [Exception]{
   echo $_.Exception|format-list -force
}

