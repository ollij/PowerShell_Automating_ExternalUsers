# Error Handling Demo

# WRITING LOG TO FILES
Function Write-Log  {
    param(
        [Parameter(Mandatory=$true,Position=1)][String]$Message,
        [string]$ForegroundColor,                
		[switch]$NoLog,
        [switch]$NoHost,        
        [switch]$WriteErrorLogfile
    )

    $logFilename = $MyInvocation.ScriptName+".log" # normal log file
    $errorLogFilename = $MyInvocation.ScriptName+".errors.log" # separate log files to use with -WriteErrorLogFile

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
        Write-Host $Message -ForegroundColor $logColor
    }

    if ($NoLog -eq $false)
    {
        $msg | Out-File -FilePath $logFilename -Encoding utf8 -Append
    }

    if ($WriteErrorLogfile) 
    {
        $msg | Out-File -FilePath $errorLogFilename -Encoding utf8 -Append 
    }
}

Function Get-NowString 
{
    return (Get-Date).ToString("yyyy-MM-dd hh:mm:ss")
}


# Bad practice: not to handle errors

function DoStuff-WithoutErrorHandling {
    $myFile = Get-Item ".\LoggingTips.ps1"
    $message = "File.FullName = "+$myFile.FullName
    Write-Log $message -ForegroundColor Green

    $myFile = Get-Item ".\nofilewiththisname.txt"
    $message = "File.FullName = "+$myFile.FullName
    Write-Host $message -ForegroundColor Green
}

DoStuff-WithoutErrorHandling # you will not notice the error on script

# Good practice

function Test-BetterPractice
{
    try 
    {
        $myFile = Get-Item ".\nosuchfile.txt" -ErrorAction Stop
        $myFile.FullName
        $message = "File.FullName = "+$myFile.FullName
        Write-Log $message -ForegroundColor Green
    } 
    catch 
    {
        Write-Log $_.Exception.Message -ForegroundColor Red -WriteErrorLogfile
        Write-Log "Aborting"
    }
}

Test-BetterPractice
