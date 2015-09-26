#requires -Version 2
function Get-IPaddress ()
{
    param (
        [parameter(Mandatory = $true,Position = 1,HelpMessage = 'Please provide a valid HOSTNAME')]
        [string]$hostname
    ) 
    <#
            .SYNOPSIS 
            Takes HOSTNAME (DNS Name) as input and does a reverse DNS lookup on that HOSTNAME.  This function returns the IP address(es) associated with it. 

            .DESCRIPTION
            Takes a HOSTNAME (DNS Name) and does a reverse DNS lookup on that HOSTNAME
            Returns IP Address(es) associated with that HOSTNAME

            .PARAMETER hostname
            Specifices the specific HOSTNAME (DNS Name) 
   
            .EXAMPLE
            C:\PS> Get-IPAddress -hostname wix.com

    #>

    $ipaddresses = [System.Net.Dns]::GetHostAddresses("$hostname").IPAddressToString

    return $ipaddresses
}
