#requires -Version 2
function Get-URLFromMessage
{
    param (
        [parameter(Mandatory = $true,Position = 1,HelpMessage = 'Please provide a .MSG file to Parse')]
        $inputtext
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
 

    Get-ChildItem $inputtext|
    ForEach-Object -Process {
        $outlook = New-Object -ComObject outlook.application
        $msg = $outlook.CreateItemFromTemplate($_.FullName)
        #$msg | Select-Object -Property *
        $url = $msg |
        Select-Object -Property body |
        Select-String -Pattern '(?:(?:https?|ftp|file)://|www\.|ftp\.)(?:\([-A-Z0-9+&@#/%=~_|$?!:,.]*\)|[-A-Z0-9+&@#/%=~_|$?!:,.])*(?:\([-A-Z0-9+&@#/%=~_|$?!:,.]*\)|[A-Z0-9+&@#/%=~_|$])' |
        ForEach-Object -Process {
            $_.Matches
        } |
        ForEach-Object -Process {
            $_.Value
        } 
    }

    Write-Host 'url: ' $url
    return $url.trim('<','>')
}
