#requires -Version 1
Function Write-LogEntry 
{
    param (
        [string]$type,
        [string]$message,
        [string]$Folder,
        [string]$CustomMessage
    )

    [bool]$CustomMessage

    $mutex = New-Object -TypeName 'Threading.Mutex' -ArgumentList $false, 'MyInterprocMutex'

    switch ($type){
        'Error' 
        {
            $mutex.waitone()
            "$((Get-Date).ToString('yyyyMMddThhmmss')) [ERROR]: $message" >> "$($Folder)\log.log"
            if ($CustomMessage)
            {
                "$((Get-Date).ToString('yyyyMMddThhmmss')) [CUSTOM MESSAGE]: $CustomMessage" >> "$($Folder)\log.log"
            }
            $mutex.ReleaseMutex()
        }
        'Info' 
        {
            $mutex.waitone()
            "$((Get-Date).ToString('yyyyMMddThhmmss')) [INFO]: $message" >> "$($Folder)\log.log"
            if ($CustomMessage)
            {
                "$((Get-Date).ToString('yyyyMMddThhmmss')) [CUSTOM MESSAGE]: $CustomMessage" >> "$($Folder)\log.log"
            }
            $mutex.ReleaseMutex()
        }
        'Debug' 
        {
            $mutex.waitone()
            "$((Get-Date).ToString('yyyyMMddThhmmss')) [DEBUG]: $message" >> "$($Folder)\log.log"
            if ($CustomMessage)
            {
                "$((Get-Date).ToString('yyyyMMddThhmmss')) [CUSTOM MESSAGE]: $CustomMessage" >> "$($Folder)\log.log"
            }
            $mutex.ReleaseMutex()
        }
    }
}
