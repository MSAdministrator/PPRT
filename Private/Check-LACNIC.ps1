#requires -Version 3
function Check-LACNIC ()
{
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true,Position = 1,HelpMessage = 'Please provide a IP address')]
        $ipaddress
    ) 
    <#
            .SYNOPSIS 
            Takes IP address as input and queries LACNIC's RDAP implementation for the IP addresses abuse contact email - RESTFul API 

            .DESCRIPTION
            Takes a IP Address and searches for LACNIC's RDAP abuse contact email for that IP based on registration data
            Returns this contact email address

            .PARAMETER ipaddress
            Specifices the specific IP address belonging to LACNIC
   
            .EXAMPLE
            C:\PS> Check-LACNIC -ipaddress '190.42.65.82'   

    #>


    $regx = "[a-z0-9!#\$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#\$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?"

    $rawdata = Invoke-WebRequest -Uri "http://rdap.lacnic.net/rdap/ip/$ipaddress" | ConvertFrom-Json

    for($i = 0;$i -lt ($rawdata.entities.vcardArray).count; $i++)
    {
        foreach ($item in $rawdata.entities.vcardArray.SyncRoot[$i])
        {
            [array]$result += $item
        }
    }
    $parsedresult = $result | Select-String -Pattern $regx
    Write-Debug -Message 'parsed result: ' $parsedresult
    if ($parsedresult.count -gt 0)
    {
        return $parsedresult
    }
    return 'NO POC FOR LACNIC'
}
