#requires -Version 2
function Get-ParsedURL ()
{
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true,
        HelpMessage = 'Please provide a URL to parse')]
        $url,

        [parameter(Mandatory = $true,
        HelpMessage = 'Please provide a log path.')]
        $logpath
    ) 
    <#
            .SYNOPSIS 
            Takes a URL as input and splits the URL down to just the hostname.  This function returns parsed URL 

            .DESCRIPTION
            Takes a URL as input and splits the URL down to just the hostname.
            This function returns parsed URL

            .PARAMETER url
            Specifices the specific URL
   
            .EXAMPLE
            C:\PS> Get-ParsedURL -url 'http://outlookadminmailaccess.bravesites.com/'

    #>

    $ReturnObject = @()

    $regexipv4 = '\b((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3} (25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b'

    if ($url -like $regexipv4)
    {
        $log = Write-LogEntry -type 'INFO' -message "URL is a IP Address - URL = $url" -Folder $logpath
        return $url
    }

    $ParsedURL = [system.net.webrequest]::Create($url)

    $AbsoluteURL = $ParsedURL.GetResponse().ResponseUri.AbsoluteUri
    $URLAuthoirty = $ParsedURL.GetResponse().ResponseUri.Authority

    $props = @{
        OriginalURL  = $url
        AbsoluteURL  = $AbsoluteURL
        URLAuthority = $URLAuthoirty
    }

    $ReturnObject = New-Object -TypeName PSObject -Property $props
    
    return $ReturnObject
}
