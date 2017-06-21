# PowerShell Logging Demo

# SIMPLE LOGGING TO HOST
Write-Host "We are using Write-Host to write a log entry." -ForegroundColor Gray

# MORE COMPLEX HOST LOGGING
function w { PARAM($p1,$p2,$p3,$p4) Write-Host $p1 $p2 $p3 $p4 }

$variable = New-Object System.DateTime(2017,6,22)
w "The value of variable is:" $variable " and we are cool with it."

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

Write-Log "Message" -ForegroundColor Red 
Write-Log "Another message" -WriteErrorLogfile

# Writing to Event-Log
# Must run once with elevated permissions (run as admin powershell window) to create event source
New-EventLog -LogName Application -Source "Engage Demo" 

Write-EventLog -LogName Application -Source "Engage Demo" -EntryType Warning -EventId 1001 -Message "Message"
