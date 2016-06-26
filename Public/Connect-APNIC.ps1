#requires -Version 3
function Connect-APNIC ()
{
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true,Position = 1,HelpMessage = 'Please provide a IP address')]
        $ipaddress
    ) 
    <#
            .SYNOPSIS 
            Takes IP address as input and queries APNIC's RDAP implementation for the IP addresses abuse contact email - RESTFul API 

            .DESCRIPTION
            Takes a IP Address and searches for APNIC's RDAP abuse contact email for that IP based on registration data
            Returns this contact email address

            .PARAMETER ipaddress
            Specifices the specific IP address belonging to APNIC
   
            .EXAMPLE
            C:\PS> Check-APNIC -ipaddress '150.42.65.82'

    #>

    $regx2 = "[a-z0-9!#\$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#\$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?"
 
    $rawdata = Invoke-WebRequest -Uri "http://rdap.apnic.net/ip/$ipaddress" | ConvertFrom-Json

    for($i = 0;$i -lt ($rawdata.entities.vcardArray).count; $i++)
    {
        foreach ($item in $rawdata.entities.vcardArray.SyncRoot[$i])
        {
            [array]$result += $item
        }
    }

    $result | Select-String -Pattern $regx2

    return $result
}
