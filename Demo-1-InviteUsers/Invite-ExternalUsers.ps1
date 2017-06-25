$startErrorCount = $Error.Count
Write-Host "Invite-ExternalUsers starts. Loading Functions.ps1..." -NoNewline
Set-Location $PSScriptRoot
. .\Functions.ps1
Write-Log " Done." -NoLog 

Write-Log "Reading config file..."

if ($PSScriptRoot) {
    $configXmlPath = $PSScriptRoot+"\ExternalUsers.config.xml"
} else {
    $configXmlPath = ".\ExternalUsers.config.xml"
}
[xml]$configXml = Get-Content -Path $configXmlPath -ErrorAction Stop

$credentials = Get-StoredCredential -Target $configXml.config.storedCredential

Write-Log "Connectiong Azure AD..."

$connection = Connect-AzureAD -Credential $credentials -TenantId $configXml.config.tenantId

Write-Log "Reading configuration and setting up the message..."

$inviteRedirectUrl = $configXml.config.inviteRedirectUrl
$adGroupToAddExternalUser = $configXml.config.adGroupForExternalUsers

$messageInfo = New-Object Microsoft.Open.MSGraph.Model.InvitedUserMessageInfo
$messageInfo.CustomizedMessageBody = $configXml.config.customizedMessageBody
$recipient = New-Object Microsoft.Open.MSGraph.Model.Recipient
$emailAddress = New-Object Microsoft.Open.MSGraph.Model.EmailAddress
$emailAddress.Address = $configXml.config.adminAccount
$recipient.EmailAddress = $emailAddress
$messageInfo.CcRecipients = $recipient

Write-Log "Reading security group information and listing the current members..."

$securityGroup = Get-AzureADGroup -SearchString $adGroupToAddExternalUser
$members = Get-AzureADGroupMember -ObjectId $securityGroup.ObjectId -all $true

if (($members -eq $null) -or ($members.count -eq $null) -or ($members.count -le 0)) {
            $ErrorMessage = "Error: 'members' is null, 'members.count' is null or 'members.count' <= 0"            
            Write-Log $ErrorMessage
            Write-Error $ErrorMessage
            Write-Log "Exiting"
            Write-Error "Exiting"
            exit                        
}

$message = "Members.count = "+$members.Count
Write-Log $message

Write-Log "Reading CSV-files..."
$csvFiles = Get-Item ".\*.csv"
if ($csvFiles -eq $null) {
    Write-Log "No CSV-files were found."
}


for ($n = 0; $n -lt $csvFiles.Count; $n++) {
    $users = Import-Csv $csvFiles[$n].FullName
    $message = "Iterating through all users in CSV file: "+$csvFiles[$n].FullName
    Write-Log $message
    for ($i=0; $i -lt $users.Count; $i++) {        
        try {
            $userName = $users[$i].DisplayName
            $email = $users[$i].Email
            $emailLower = $email.ToLower()
            $found = $false
            for ($j=0; $j -lt $members.Count; $j++) {
                if ($members[$j].Mail -ne $null) {
                    if ($members[$j].Mail.ToLower() -eq $emailLower) {
                        $found = $true
                        break
                    }
                }
            }
            if ($found) {
                $msg =  "User already found: "+$email+" Skipping."
                Write-Log $msg
            } else {
                $msg =  "Inviting user: "+$email
                Write-Log $msg
                $invitation = New-AzureADMSInvitation -InvitedUserDisplayName $userName -InvitedUserEmailAddress $email -SendInvitationMessage $true -InvitedUserMessageInfo $messageInfo -InviteRedirectUrl $inviteRedirectUrl 
                $msg =  " Adding to security group... invitation.InvitedUser.Id:"+$invitation.InvitedUser.Id
                Write-Log $msg
                Add-AzureADGroupMember -ObjectId $securityGroup.ObjectId -RefObjectId $invitation.InvitedUser.Id
            }
        } catch {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Error $ErrorMessage
            Write-Error $FailedItem
        }
    }
    Write-Log "Moving file to csv_archive folder..."
    if (!(Test-Path .\csv_archive)) {
        New-Item csv_archive -ItemType Directory
    }    
    Move-Item $csvFiles[$n].FullName .\csv_archive
}
