#requires -Version 2
function Get-URLFromMessage
{
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true,
        HelpMessage = 'Please provide a .MSG file.')]
        [PSTypeName('PPRT.Message')]
        $Message,

        [Parameter(Mandatory = $true)]
        [ValidateScript({ if (Test-Path $_){$true}else{ throw 'Please provide a valid path for LogPath' }})]
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

    Begin
    {
        $URLPattern = '(?:(?:https?|ftp|file)://|www\.|ftp\.)(?:\([-A-Z0-9+&@#/%=~_|$?!:,.]*\)|[-A-Z0-9+&@#/%=~_|$?!:,.])*(?:\([-A-Z0-9+&@#/%=~_|$?!:,.]*\)|[A-Z0-9+&@#/%=~_|$])'
 
        $URLObject = @()
    }
    Process
    {
        $log = Write-LogEntry -type Info -message "Get-URLFromMessage: Getting URL from $($Message.Subject)" -Folder $LogPath

        $URL = $Message.body | Select-String -AllMatches $URLPattern | Select-Object -ExpandProperty Matches | Select-Object -ExpandProperty Value

        $props = @{
            URL  = $URL
            Name = $Message.Subject
        }

        $URLObject = New-Object -TypeName PSObject -Property $props

        $log = Write-LogEntry -type Info -message 'Get-URLFromMessage: Getting URL complete!' -Folder $LogPath

        Add-ObjectDetail -InputObject $URLObject -TypeName PPRT.PhishingURL 
    }
    End
    {

    }
}
