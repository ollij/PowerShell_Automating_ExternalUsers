# Functions.ps1

function ConvertTo-B2BCsv {
    PARAM($UserObjects,$FilenamePart,[int]$LinesPerFile)
    $fileCreated = $false
    [int]$fileNumber = 1
    [int]$linesWritten = 0
    $csvFilePath = ""    
    $message = " Writing "+$linesPerFile+" lines per file." 
    Write-Log $message
    foreach ($userObject in $UserObjects) {        
        if ($fileCreated -eq $false) {
            $csvFilePath = $PSScriptRoot+"\"+$FilenamePart+$fileNumber+".csv"
            if (Test-Path $csvFilePath) {
                Write-Log " Removing existing file " 
                Write-Log $csvFilePath
                Remove-Item -Path $csvFilePath
            }    
            Add-Content -Path $csvFilePath -Value $configXml.config.csv.header -Encoding UTF8        
            Write-Log " New File created: " 
            Write-Log $csvFilePath
            $fileCreated = $true
        }
        $line = $userObject.UserPrincipalName+","+$userObject.DisplayName+","+$configXml.config.csv.invitationText+","+$configXml.config.csv.inviteRedirectUrl+","+$configXml.config.csv.invitedToApplications+","+$configXml.config.csv.invitedToGroups+","+$configXml.config.csv.ccEmailAddress+","+$configXml.config.csv.language
        Add-Content -Path $csvFilePath -Value $line -Encoding UTF8
        $linesWritten++
        if ($linesWritten -eq $LinesPerFile) {
            $linesWritten = 0
            $fileNumber++
            $fileCreated = $false            
        }
    }
}


function ExcludeBy-UserType {
    PARAM($UserObjects, $TextValue)
    $newUserObjects1 = @()    
    foreach ($userObject in $userObjects) {        
        [string]$thisUserType = New-Object String($userObject.UserType.ToString())                   
        if ($thisUserType -eq $TextValue[0]) {
            $message = "  Excluding "+$UserObject.UserPrincipalName
            Write-Log $message            
        } else {                        
            $newUserObjects1 += $userObject
        }
    }    
    return $newUserObjects1
}
function ExcludeBy-UserPrincipalName {
    PARAM($UserObjects, [string]$ContainsAny)
    $newUserObjects2 = @()
    $split = $ContainsAny.ToLower().Split(';')    
    foreach ($userObject in $userObjects) {        
        $excludeThis = $false
        foreach ($c in $split) {              
            $x = $c.Trim()
            if ($UserObject.UserPrincipalName.ToLower().Contains($x)) {
                $message = "  Excluding "+$UserObject.UserPrincipalName+" based on exclusion rule: contains('"+$x+"')"
                Write-Log $message            
                $excludeThis = $true                
            }             
        }
        if ($excludeThis -eq $false) {            
            $newUserObjects2 += $userObject
        }
    }    
    return $newUserObjects2
}


Function Write-Log  
{
    param(
        [Parameter(Mandatory=$true,Position=1)][String]$Message,
        [string]$ForegroundColor,                
		[switch]$NoNewLine,
        [switch]$NoLog,
        [switch]$NoHost,        
        [switch]$WriteErrorLogfile
    )

    $logFilename = $MyInvocation.ScriptName+".log"
    $errorLogFilename = $MyInvocation.ScriptName+".errors.log"

    $msg = (Get-NowString)+" "+$Message

    if ([String]::IsNullOrEmpty($ForegroundColor))
    {
        $logColor = 'White'
    }
    else 
    {
        $logColor = $ForegroundColor    
    }
    

    if ($NoHost -eq $false)
    {
        if ($NoNewLine)
        {
            Write-Host $Message  -ForegroundColor $logColor
        }
        else 
        {            
            Write-Host $Message -ForegroundColor $logColor
        }
    }

    if ($NoLog -eq $false)
    {
        if ($NoNewLine)
        {
            $msg | Out-File -FilePath $logFilename -Encoding utf8 -Append 
        }
        else 
        {            
            $msg | Out-File -FilePath $logFilename -Encoding utf8 -Append 
        }
    }

    if ($WriteErrorLogfile) 
    {
        if ($NoNewLine)
        {
            $msg | Out-File -FilePath $errorLogFilename -Encoding utf8 -Append 
        }
        else 
        {            
            $msg | Out-File -FilePath $errorLogFilename -Encoding utf8 -Append 
        }
    }
}

Function Get-NowString 
{
    return (Get-Date).ToString("yyyy-MM-dd hh:mm:ss")
}

Function Get-TodayString
{
    return (Get-Date).ToString("yyyy-MM-dd")
}

Function Clear-Logfiles
{
    $nowString = Get-NowString        
    $logFilename = $MyInvocation.ScriptName+".log"
    if (Test-Path $logFilename)
    {        
        $newName = $MyInvocation.ScriptName+" "+$nowString+ ".log" 
        Rename-Item $logFilename -NewName $newName 
    }    
    $errorLogFilename = $MyInvocation.ScriptName+".errors.log"
    if (Test-Path $errorLogFilename)
    {
        $newName = $MyInvocation.ScriptName+" "+$nowString+ ".errors.log" 
        Rename-Item $logFilename -NewName $newName 
    }    
}