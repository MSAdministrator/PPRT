#requires -Version 2
function Get-AbsoluteUri
{
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true,
        HelpMessage = 'Please provide a .MSG file.')]
        [PSTypeName('PPRT.PhishingURL')]
        $URLObject
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

    if ($URLObject | Select-Object -Property URL)
    {
        $url = $URLObject.URL -as [system.URI]

        if (!($url.AbsoluteURI -ne $null -and $url.Scheme -match '[http|https]'))
        {
            Write-Error -Message 'URL is not formatted correctly'
        }

        $encodedurl = [system.net.webrequest]::Create($url)
    }
    else
    {
        $url = $URLObject.URL -as [system.URI]

        if (!($url.AbsoluteURI -ne $null -and $url.Scheme -match '[http|https]'))
        {
            Write-Error -Message 'URL is not formatted correctly'
        }

        $encodedurl = [system.net.webrequest]::Create($url)
    }
    
    $AbsoluteURL = $encodedurl.GetResponse().ResponseUri.AbsoluteUri
    $URLAuthority = $encodedurl.GetResponse().ResponseUri.Authority

    Import-Clixml -Path "$(Split-Path -Path $Script:MyInvocation.MyCommand.Path)\Private\shorturls.xml" | ForEach-Object -Process {
        if ($AbsoluteURL -like '$_')
        {
            #call Get-LongUrl to call API to resolve to the normal/long url
            $ExpandedURL = Add-ObjectDetail -InputObject $AbsoluteURL -TypeName PPRT.PhishingURL
            $FinalURL = Expand-ShortURL $ExpandedURL
        }
    }

    $props = @{
        OriginalURL  = $URLObject
        EncodedURL   = $encodedurl
        AbsoluteURL  = $AbsoluteURL
        URLAuthority = $URLAuthority
    }

    $ReturnObject = New-Object -TypeName PSObject -Property $props

    Add-ObjectDetail -InputObject $ReturnObject -TypeName PPRT.ExpandedURL
}
