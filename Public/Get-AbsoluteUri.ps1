#requires -Version 2
function Get-AbsoluteUri
{
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true,
        HelpMessage = 'Please provide a .MSG file.')]
        #[PSTypeName('PPRT.PhishingURL')]
        $URLObject,

        [Parameter(Mandatory = $true)]
        [ValidateScript({ if (Test-Path $_){$true}else{ throw 'Please provide a valid path for LogPath' }})]
        $LogPath
    ) 
    <#
            .SYNOPSIS 
            This function will get the Absolute Uri for a given URLObject

            .DESCRIPTION
            This function checks and verifies that URL passed in is formed for the correct scheme
            We will get the response from contacting the website and return detailed information about
            the given URL.

            .PARAMETER URLObject
            A PPRT.PhishingURL Object Type that will checked and more detail returned.

            .PARAMETER LogPath 
            This parameter is a folder path that you want to log to

            .EXAMPLE
            C:\PS> Get-AbsoluteUri -URLObject $URL -LogPath $LogPath

    #>

    $ReturnObject = @()
    $AbsoluteURL = @()
    $URLAuthority = @()

    Add-Type -AssemblyName System.Web

    $TempURL = $URLObject.URL

    if (($null -eq $TempURL.AbsoluteURI -and $TempURL.Scheme -match '[http|https]'))
    {
        $log = Write-LogEntry -type Error -message 'Get-AbsoluteUri: URL is not the correct scheme' -Folder $LogPath
    }

    $log = Write-LogEntry -type Info -message 'Get-AbsoluteUri: Creating WebRequest' -Folder $LogPath
       

    $encodedurl = [system.net.webrequest]::Create($($TempURL))
    $Response = $null
    
    try
    {
        $log = Write-LogEntry -type Info -message 'Get-AbsoluteUri: Getting Response' -Folder $LogPath
        $Response = $encodedurl.GetResponse()

        $AbsoluteURL = $Response.ResponseUri.AbsoluteUri
        $URLAuthority = $Response.ResponseUri.Authority
    }
    catch
    {
        $log = Write-LogEntry -type Error -message 'Get-AbsoluteUri: ERROR GETTING RESPONSE!!!' -Folder $LogPath -CustomMessage 'BREAK!'
        continue
    }
    finally
    {
        # Clear the response, otherwise the next HttpWebRequest may fail... (don't know why)
        if ($Response -ne $null) 
        {
            $Response.Close()
        }  
    }

    try
    {
        Import-Clixml -Path "$(Split-Path -Path $Script:MyInvocation.MyCommand.Path)\Private\shorturls.xml" | ForEach-Object -Process {
            if ($AbsoluteURL -like '$_')
            {
                $log = Write-LogEntry -type Info -message "Get-AbsoluteUri: URL matches a Short URL - $AbsoluteURL = $($_)" -Folder $LogPath

                #call Get-LongUrl to call API to resolve to the normal/long url
                $ExpandedURL = Add-ObjectDetail -InputObject $AbsoluteURL -TypeName PPRT.PhishingURL

                $log = Write-LogEntry -type Info -message 'Get-AbsoluteUri: Calling Get-AbsoluteUri Again' -Folder $LogPath

                $FinalURL = Get-AbsoluteUri $ExpandedURL
            }
        }
    }
    catch
    {
        $log = Write-LogEntry -type Error -message 'Get-AbsoluteUri: Unable to find shorturls.xml!' -Folder $LogPath
    }

    $props = @{
        OriginalURL  = $URLObject
        EncodedURL   = $($encodedurl)
        AbsoluteURL  = $AbsoluteURL
        URLAuthority = $URLAuthority
    }

    $ReturnObject = New-Object -TypeName PSObject -Property $props

    $log = Write-LogEntry -type Info -message 'Get-AbsoluteUri: Completed Successfully!' -Folder $LogPath

    Add-ObjectDetail -InputObject $ReturnObject -TypeName PPRT.Uri
}
