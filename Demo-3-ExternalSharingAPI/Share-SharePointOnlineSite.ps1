$siteUrl = "https://yourtenant.sharepoint.com/sites/site"

$externalUserEmails = "user1@somedomain.com", "user3@somedomain.com",	"user3@somedomain.com"

$emailBody = "Hello!<br/><br/>Please try to access the site with your <i>name@somedomain.com</i> account."

# Change the path to valid location on your machine
add-type -Path "C:\Users\olli\Source\Repos\PnP2\Samples\Core.ExternalSharing\packages\Microsoft.SharePointOnline.CSOM.16.1.3912.1204\lib\net45\Microsoft.SharePoint.Client.dll"
add-type -Path "C:\Users\olli\Source\Repos\PnP2\Samples\Core.ExternalSharing\packages\Microsoft.SharePointOnline.CSOM.16.1.3912.1204\lib\net45\Microsoft.SharePoint.Client.Runtime.dll"
$context = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
$ss = Read-Host -Prompt "Enter password" -AsSecureString
$context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials("admin@user1@somedomain.com.onmicrosoft.com", $ss)
$web = $context.Web
$context.Load($web)
$context.Load($web.AssociatedMemberGroup)
$context.ExecuteQuery()

$userJsonTemplate = '
    "Key" : "{0}", 
    "Description" : "{0}", 
    "DisplayText" : "{0}", 
    "EntityType": "", 
    "ProviderDisplayName" : "", 
    "ProviderName" : "", 
    "IsResolved" : true, 
    "EntityData" : {
        "SPUserID" : "{0}", 
        "Email" : "{0}", 
        "IsBlocked" : "False", 
        "PrincipalType" : "UNVALIDATED_EMAIL_ADDRESS", 
        "AccountName" : "{0}", 
        "SIPAddress" : "{0}"
    }, 
    "MultipleMatches" : []
'

$jsonUsers = '['
$userAlreadyAdded = $false;
foreach ($email in $externalUserEmails) {
    if ($userAlreadyAdded) {
        $jsonUsers += ","
    }
    $s = $userJsonTemplate.Replace("{0}", $email)
    $jsonUsers += "{" +  $s + "}"
    $userAlreadyAdded = $true
}
$jsonUsers += "]"

#$jsonUsers

$propageAcl = $false
$sendEmail = $true
$includedAnonymousLinkInEmail = $false
$emailSubject = ""

$roleValue = "group:"+$web.AssociatedMemberGroup.Id

# https://msdn.microsoft.com/en-us/library/office/mt684216.aspx for the documentation of ShareObjec
[Microsoft.SharePoint.Client.SharingResult]$result = [Microsoft.SharePoint.Client.Web]::ShareObject( `
    $web.Context, $web.Url, $jsonUsers, $roleValue, `
    $web.AssociatedMemberGroup.Id, $propageAcl, `
    $sendEmail, $includedAnonymousLinkInEmail, $emailSubject, $emailBody) 
$context.Load($result)
$context.ExecuteQuery()
