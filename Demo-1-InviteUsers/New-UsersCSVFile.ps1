$startErrorCount = $Error.Count
Write-Host "New-UsersCSVFile starts. Loading Functions.ps1..." -NoNewline
Set-Location $PSScriptRoot
. .\Functions.ps1
Write-Log " Done." -NoLog 

$configXmlPath = ".\UsersCSVFile.config.xml"
[xml]$configXml = Get-Content -Path $configXmlPath -ErrorAction Stop

$credentials = Get-StoredCredential -Target $configXml.config.storedCredential

Write-Log "Connecting MsolService..."
Connect-MsolService -Credential $credentials

Write-Log "Handling include filters..."
if ($configXml.config.includeFilters.allUsers.ToString() -eq "true") {
    Write-Log " Reading all users..."
    $users = Get-MsolUser -All
} else {
    $users = @()
    $groups = MSOnlineExtended\Get-MsolGroup -All
    foreach($item in $configXml.config.includeFilters.groups.ChildNodes) {          
        $message= " Adding users in group: "+$item.InnerText
        Write-Log $message
        $g = $groups | Where-Object { $_.DisplayName -eq $item.InnerText }
        if ($g -eq $null) {
            $message = "  includeFilters/groups - Group not found: "+$item.InnerText
            Write-Error $message
        } else {             
            $groupMembers = Get-MsolGroupMember -GroupObjectId $g.ObjectId                    
            foreach ($groupMember in $groupMembers) {
                if ($groupMember.GroupMemberType -eq "User") {
                    $userAlreadyAdded = $false
                    foreach($user in $users) {
                        if ($user.ObjectId -eq $groupMember.ObjectId) {
                                $userAlreadyAdded = $true
                                break
                        }
                    }
                    if ($userAlreadyAdded -eq $false ) {
                        $users += Get-MsolUser -ObjectId $groupMember.ObjectId                    
                        $message = "  Member added: "+$groupMember.DisplayName
                        Write-Log $message
                    }
                } 
            }
        }
    }

}
$userCount = $users.Count
$message = "Usercount: "+$userCount
Write-Log $message

Write-Log "Handling exclude filters..."

foreach ($item in $configXml.config.excludeFilters) {
    $property = $item.excludeFilter.property    
    if ($property -eq "UserType") {
        $value = $item.excludeFilter.value
        $message = " Excluding by UserType: "+$value
        Write-Log $message        
        $users = ExcludeBy-UserType -UserObjects $users -TextValue $value
        $userCount = $users.Count
        $message = " Usercount: "+$userCount
        Write-Log $message
    }
    if ($property -eq "UserPrincipalName") {
        Write-Log " Excluding by UserPrincipalName"
        [string]$s = $item.excludeFilter.containsAny
        if ([string]::IsNullOrEmpty($s) -eq $false) {        
            $s = $s.ToLower()
            $users = ExcludeBy-UserPrincipalName -UserObjects $users -ContainsAny $s

        }
        $userCount = $users.Count
        $message = " Usercount: "+$userCount
        Write-Log $message
    }
}

Write-Log "Creating CSV file(s)..."
$today = Get-TodayString
$random = -join ((97..122) | Get-Random -Count 5 | % {[char]$_})
$fnPart = "users-"+$today+"-"+$random 
ConvertTo-B2BCsv -UserObjects $users -FilenamePart $fnPart -LinesPerFile $configXml.config.csv.linesPerFile

Write-Log "Checking for errors..."
$endErrorCount = $Error.Count
$errors = $endErrorCount - $startErrorCount
$message = "Errors: "+$Errors
Write-Log $message
for ($i=0; $i -lt $errors; $i++) {
    $e = $Error[$i+$startErrorCount]
    $message = "Error: "+($i+1)
    Write-Log $message -WriteErrorLogfile
    $message = " Exception: "+$e.Exception
    Write-Log $message -WriteErrorLogfile
    $message = " ErrorDetails: "+$e.ErrorDetails
    Write-Log $message -WriteErrorLogfile
    $message = " InvocationInfo: "+$e.InvocationInfo    
    Write-Log $message -WriteErrorLogfile
    $message = " ScriptStackTrace: "+$e.ScriptStackTrace
    Write-Log $message -WriteErrorLogfile
}