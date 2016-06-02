#requires -Version 3
function Get-LongUrl ()
{
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true,Position = 1,HelpMessage = 'Please provide a short URL')]
        $shorturl,

        [parameter(Mandatory = $true,
        HelpMessage = 'Please provide a log path.')]
        $logpath
    ) 
    <#
        .SYNOPSIS 
        Takes a short url and converts it to a long url

        .DESCRIPTION
        This function takes a tinyurl and converts it to a long/normal URL using the System.Net.WebRequest class.
        This function will continue to call itself until the URL has been expanded successfully.

        .PARAMETER shorturl   
        This parameter needs to be a TinyUrl at this time.

        .PARAMETER logpath   
        This parameter is a folder path that you want to log to

        .EXAMPLE
        C:\PS> Get-LongUrl -shorturl $shortUrlVariable -logpath $logpath

    #>

    Add-Type -AssemblyName System.Web

    $encodedurl = [system.net.webrequest]::Create($shorturl)
    
    $ExpandedURL = $encodedurl.GetResponse().ResponseUri.AbsoluteUri
    $URLAuthority = $encodedurl.GetResponse().ResponseUri.Authority

    Import-Clixml -Path "$(Split-Path $Script:MyInvocation.MyCommand.Path)\Private\shorturls.xml" | ForEach-Object {
        if ($ExpandedURL -like '$_')
        {
            #call Get-LongUrl to call API to resolve to the normal/long url
            $longurl = Get-LongUrl $ExpandedURL
            Write-LogEntry -type 'INFO' -message "URL is a shortened URL - Expanded URL = $ExpandedURL" -Folder "$logpath\PPRTLog.txt"
        }
    }

    $props = @{
        OriginalURL = $shorturl
        AbsoluteURL = $ExpandedURL
        URLAuthority = $URLAuthority
    }

    $ReturnObject = New-Object -TypeName PSObject -Property $props

    return $ReturnObject
}
