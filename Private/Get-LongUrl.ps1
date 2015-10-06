#requires -Version 3
function Get-LongUrl ()
{
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true,Position = 1,HelpMessage = 'Please provide a short URL')]
        $shorturl
    ) 
    <#
            .SYNOPSIS 
            Takes a short url and converts it to a long url

            .DESCRIPTION
            This function takes a tinyurl and converts it to a long/normal URL using http://api.longurl.org RESTFul API.

            .PARAMETER shorturl   
            This parameter needs to be a TinyUrl at this time.

            .INPUTS

            .OUTPUTS

            .EXAMPLE
            C:\PS> Get-LongUrl -shorturl $shortUrlVariable  

    #>

    Add-Type -AssemblyName System.Web
    $encodedurl = [System.Web.HttpUtility]::UrlEncode($shorturl)
    $longurl = Invoke-WebRequest -Uri "http://api.longurl.org/v2/expand?url=$encodedurl&format=json"
    $convertedurl = $longurl.content | ConvertFrom-Json
    return $convertedurl.'long-url'
}
