#requires -Version 2
function Get-URLFromMessage
{
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true,
        HelpMessage = 'Please provide a .MSG file.')]
        [PSTypeName('PPRT.Message')]
        $MessageObject,

        [Parameter(Mandatory = $true)]
        $LogPath
    ) 
    <#
            .SYNOPSIS 
            Takes a .MSG file and parses the links from the message. This function returns the full URL within an email. 

            .DESCRIPTION
            Takes a .MSG file and parses the links from the message.
            This function returns the full URL within an email. 

            .PARAMETER inputtext
            Specifices the specific .MSG to parse
   
            .EXAMPLE
            C:\PS> Get-URLFromMessage 'C:\Users\UserName\Desktop\PHISING_EMAILS\Dear Email User.msg'

    #>
 
    $urlobject = @()

    foreach ($msg in $MessageObject)
    { 
        $log = Write-LogEntry -type Info -message "Get-URLFromMessage: Getting URL from $($msg.Subject)" -Folder $LogPath

        $Subject = $msg.Subject
        $msg.body |
        Select-String -Pattern '(?:(?:https?|ftp|file)://|www\.|ftp\.)(?:\([-A-Z0-9+&@#/%=~_|$?!:,.]*\)|[-A-Z0-9+&@#/%=~_|$?!:,.])*(?:\([-A-Z0-9+&@#/%=~_|$?!:,.]*\)|[A-Z0-9+&@#/%=~_|$])' |
        `
        ForEach-Object -Process {
            $_.Matches
        } |
        ForEach-Object -Process {
            $URL = $_.Value

            $props = @{
                URL  = ($_.Value).trim('<','>')
                Name = $Subject
            }

            $urlobject = New-Object -TypeName PSObject -Property $props

            $log = Write-LogEntry -type Info -message 'Get-URLFromMessage: Getting URL complete!' -Folder $LogPath

            Add-ObjectDetail -InputObject $urlobject -TypeName PPRT.PhishingURL 
        }
    }
}
