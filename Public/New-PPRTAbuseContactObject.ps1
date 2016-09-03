#requires -Version 2
<#
        .Synopsis
        New-PPRTAbuseContactObject is a function to gather information about the abuse/registration POC
        .DESCRIPTION
        New-PPRTAbuseContactObject is a function to gather information about the abuse/registration POC.
        Additionally, this function will:
        Check to see if the URL is alive
        Get detailed URL information (AbsoluteUri)
        Get the IP Address of the host/URL
        Find out which WHOIS owns this IP
        Run against the owner/registrars RDAP/WHOIS API
        .EXAMPLE
        New-PPRTAbuseContactObject -URLObject $URLObject -LogPath $LogPath
        .EXAMPLE
        $URLObject | New-PPRTAbuseContactObject -LogPath $LogPath
#>
function New-PPRTAbuseContactObject
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable],[String])]
    Param
    (
        [Parameter(Mandatory = $true,
                ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        # [PSTypeName('PPRT.PhishingURL')]
        $URLObject,

        [Parameter(Mandatory = $true)]
        [ValidateScript({ if (Test-Path $_){$true}else{ throw 'Please provide a valid path for LogPath' }})]
        $LogPath
    )
    Begin
    {
        $AllAttachments = @()
        $ReturnOjbect = @()
        
        $object = @{}
    }
    Process
    {
        foreach ($url in $URLObject)
        {
            $Body = @{}

            $Body.URL = $url

            $URLAlive = @()

            # try
            # {
                
            $log = Write-LogEntry -type Info -message "New-PPRTAbuseContactObject: Verifying that URL is alive - $($url)" -Folder $LogPath
            
            $log = Write-LogEntry -type Info -message 'New-PPRTAbuseContactObject: Getting AbsoluteUri' -Folder $LogPath

            $AbsoluteUri = Get-AbsoluteUri -URLObject $url -LogPath $LogPath

            if ($AbsoluteUri -eq $null)
            {
                $log = Write-LogEntry -type Error -message 'New-PPRTAbuseContactObject: AbsoluteUri is Null' -Folder $LogPath
                $AbsoluteUri = $null
            }

            $Body.AbsoluteUri = $AbsoluteUri

            $AbsoluteUri.URLAuthority

            [array]$ipaddress = Get-IPaddress -hostname $($AbsoluteUri.URLAuthority) -LogPath $LogPath
            if ($ipaddress -eq $null)
            {
                $log = Write-LogEntry -type Error -message 'New-PPRTAbuseContactObject: IPAddress is Null' -Folder $LogPath
                $ipaddress = $null
            }

            $Body.IPAddress = $ipaddress

            foreach ($ip in $ipaddress)
            {
                #based on the ipaddress we are going to get which WHOIS/RDAP to use
                $whoisdb = Get-WhichWHOIS -ipaddress $ip

                if ($whoisdb.WHOIS -eq $null )
                {
                    $log = Write-LogEntry -type Error -message 'New-PPRTAbuseContactObject: WHOIS is Null' -Folder $LogPath
                    $whoisdb = $null
                }

                $Body.WHOIS = $whoisdb
    
                #based on info from Get-WhichWHOIS we will then begin those specific API calls
                switch ($whoisdb){
                    'arin' 
                    {
                        [array]$abusecontact = Check-ARIN $ip
                    }
                    'ripe' 
                    {
                        [array]$abusecontact = Check-RIPE $ip
                    }
                    'apnic' 
                    {
                        $abusecontact = Check-APNIC $ip
                    }
                    'lacnic' 
                    {
                        [array]$abusecontact = Check-LACNIC $ip
                    }
                    'afrnic' 
                    {
                        $abusecontact = 'NOCONTACT'
                    }
                    $null 
                    {

                    }
                }

                $Body.AbuseContact = $abusecontact
            }
                

            Add-ObjectDetail -InputObject $Body -TypeName PPRT.AbuseContact
        }      
    }
    End
    {
    }
}
