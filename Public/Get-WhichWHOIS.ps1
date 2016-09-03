#requires -Version 2
function Get-WhichWHOIS ()
{
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true,Position = 1,HelpMessage = 'Please provide a valid IP Address')]
        [string]$ipaddress,

        [Parameter(Mandatory = $true)]
        [ValidateScript({ if (Test-Path $_){$true}else{ throw 'Please provide a valid path for LogPath' }})]
        $LogPath
    ) 
    <#
            .SYNOPSIS 
            Takes IPAddress as input and finds which whois should be used.

            .DESCRIPTION
            Takes an IPAddress and splits the first octect of the IP address
            Takes the first octect and compares against arrays of registrars
            Returns which whois should be used

            .PARAMETER ipaddress
            Specifies the ipdadress we are wanting information on
   
            .EXAMPLE
            C:\PS> Get-WhichWHOIS -ipaddress '189.84.54.56'

    #>

    $RegistryObject = @()

    try
    {
        $RegistryData = Invoke-RestMethod -Uri 'http://www.iana.org/assignments/ipv4-address-space/ipv4-address-space.xml'
    }
    catch
    {
        Write-LogEntry -type ERROR -message 'UNABLE TO REACH IANA.org' -Folder $LogPath
        
        throw 'Unable to reach http://www.iana.org/assignments/ipv4-address-space/ipv4-address-space.xml at this time'
    }

    for ($i = 1; $i -le $RegistryData.registry.record.Count; $i++)
    {
        if ($null -ne $RegistryData.registry.record[$i].prefix)
        {
            $TrimmedIP = $(($RegistryData.registry.record[$i].prefix).TrimStart("0"))
            
            $ComparisonIP = $TrimmedIP.Substring(0,$TrimmedIP.Length - 2)
            
            if ($($ipaddress.Split('{.}')[0]) -eq $ComparisonIP)
            {
                if ($RegistryData.registry.record[$i].whois -ne $null)
                {
                    $WHOIS = ($RegistryData.registry.record[$i].whois)
                    
                    $props = @{
                        IPAddress = $ipaddress
                        WHOIS = $(($WHOIS).Split('{.}')[1])
                        OriginalWHOIS = $WHOIS
                    }
                    
                    $TempObject = New-Object -TypeName PSCustomObject -Property $props
                    
                    Add-ObjectDetail -InputObject $TempObject -TypeName PPRT.WHOIS
                }
            }
        }
    }
}
