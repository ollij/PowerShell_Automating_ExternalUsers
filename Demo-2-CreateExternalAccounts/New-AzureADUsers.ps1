cd "C:\PowerShell\PowerShell_Automating_ExternalUsers\Demo-2-CreateExternalAccounts"

function Get-RandomPassword {
    $retVal = -join ((65..90) + (97..122) | Get-Random -Count 8 | % {[char]$_}) # 8 kIrjaInTa
    $retVal += -join ((33..38) | Get-Random -Count 1 | % {[char]$_})  # erikoismerkki
    $retVal += -join ((48..57) | Get-Random -Count 1 | % {[char]$_})  # numero
    return $retVal  
}

function Send-InvitaionMail {
    param(
        $SiteUrl, $From, $To, $NewAccount, $NewPassword, $SenderCredential
    )
    $mail = New-Object System.Net.Mail.MailMessage($From,$To)
    $mail.Subject = "Extranet account created!"
    $mail.IsBodyHtml = $true
    $mail.Body = @"
You have been created $NewAccount to Extranet.<br>
<br>
Your password is $NewPassword and you need to change it when you login for the first time.<br>
<br>
Please open $SiteUrl using your new account.
"@

    $client = New-Object System.Net.Mail.SmtpClient("smtp.office365.com", 587)
    $client.EnableSsl = $true
    $client.Credentials = $SenderCredential 
    $client.Send($mail)
}

Write-Host "Reading newusers.xlsx..."
$excelFile = New-Object -ComObject Excel.Application
$x = $excelFile.Workbooks.Open((Get-Item -Path ".\" -Verbose).FullName+"\newusers.xlsx")
$sheet = $excelFile.Sheets.Item(1)

$sharePointOnlineSiteUrl = $sheet.Cells.Item(1,2).Value2
Write-Host "SharePoint Online Site Url: "$sharePointOnlineSiteUrl
$emailFromAccount = $sheet.Cells.Item(2,2).Value2
Write-Host "Email From Account        : "$emailFromAccount 
$spGroupToAdd = $sheet.Cells.Item(3,2).Value2
Write-Host "SharePoint Group for users: "$spGroupToAdd 
$tenantAdminSite = $sheet.Cells.Item(5,2).Value2
Write-Host "Tenant admin site         : "$tenantAdminSite

$credential = Get-StoredCredential -Target "yourStoredCredentialTarget"
Write-Host "Connectiong MsolService..."
Connect-MsolService -Credential $credential
Write-Host "Connectiong SPOService..."
Connect-SPOService -Url $tenantAdminSite -Credential $credential

$rowsToProcess = $true
$row=7

while ($rowsToProcess)  {
    $firstName = $sheet.Cells.Item($row,1).Value2
    if ([string]::IsNullOrEmpty($firstName)) {
        $rowsToProcess = $false
    } else {
        $lastName = $sheet.Cells.Item($row,2).Value2
        $displayName = $sheet.Cells.Item($row,3).Value2
        $userPrincipalName = $sheet.Cells.Item($row,4).Value2
        $userType = $sheet.Cells.Item($row,5).Value2
        [string[]]$alternateEmailAddresses = $sheet.Cells.Item($row,6).Value2
        $password = Get-RandomPassword
        $user = New-MsolUser -UserPrincipalName  $userPrincipalName -FirstName $firstName -LastName $lastName -DisplayName $displayName -UserType $userType -AlternateEmailAddresses $alternateEmailAddresses -Password $password 
        if ($user -ne $null) {
            $failed = $false
            if (($spGroupToAdd -ne '') -and ($spGroupToAdd -ne $null)) {
                $currentTry = 0
                while ($true) {
                    Start-Sleep -Seconds 1
                    $failed = $false
                    $currentTry++
                    try {
                        Add-SPOUser -LoginName $userPrincipalName -Group $spGroupToAdd -Site $sharePointOnlineSiteUrl
                    } catch {
                        $failed = $true
                        Write-Host "." -NoNewline
                    }
                    if ($failed -eq $false -or $currentTry -gt 100) {
                        break;
                    }
                }
            }
            if ($failed) {
                Write-Host "$userPrincipalName could not be added to $spGroupToAdd" -ForegroundColor Red
                Write-Host "No mail to the user has been send."
            } else {
                Write-Host "$userPrincipalName created successfully, sending mail..."
                Send-InvitaionMail -SiteUrl $sharePointOnlineSiteUrl -From $emailFromAccount -To $alternateEmailAddresses -NewAccount $userPrincipalName -NewPassword $password -SenderCredential $credential
            }
        } else {
            Write-Error "Failed to create $userPrincipalName!"
        }
        $row++    
    }
}
$excelFile.Workbooks.Close()
Write-Host "All done."