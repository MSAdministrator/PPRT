#requires -Version 3
function Check-RIPE ()
{
    param (
        [parameter(Mandatory = $true,Position = 1,HelpMessage = 'Please provide a IP address')]
        $ipaddress
    ) 
    <#
            .SYNOPSIS 
            Takes IP address as input and queries RIPE's WHOIS implementation for the IP addresses abuse contact email - RESTFul API 

            .DESCRIPTION
            Takes a IP Address and searches for RIPE's WHOIS abuse contact email for that IP based on registration data
            Returns this contact email address

            .PARAMETER ipaddress
            Specifices the specific IP address belonging to RIPE
  
            .EXAMPLE
            C:\PS> Check-RIPE -ipaddress '195.42.65.82'   

    #>


    $abusecontact = Invoke-RestMethod -Uri "http://rest.db.ripe.net/abuse-contact/$ipaddress"

    $result = $abusecontact.'abuse-resources'.'abuse-contacts'.email

    return $result
}
