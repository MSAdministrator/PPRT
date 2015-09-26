#requires -Version 2
function Get-ParsedURL ()
{
    param (
        [parameter(Mandatory = $true,Position = 1,HelpMessage = 'Please provide a URL to parse')]
        [string]$url
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

    $regexipv4 = '\b((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3} (25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b'

    if ($url -like $regexipv4)
    {
        Write-Host -Object 'URL IS IP'
    }

    $parsedurl = ([System.Uri]$url).Host.split('.')[-2..-1] -join '.'
    Write-Host 'parsedurl: ' $parsedurl
    return $parsedurl
}
