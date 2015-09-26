#requires -Version 2
function Get-WhichWHOIS ()
{
    param (
        [parameter(Mandatory = $true,Position = 1,HelpMessage = 'Please provide a valid IP Address')]
        [string]$ipaddress
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

    $arinips = (23, 24, 35, 40, 45, 47, 50, 54, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 96, 97, 98, 99, 100, 104, 107, 108, 128, 129, 130, 131, 132, 134, 135, 136, 137, 138, 139, 140, 142, 143, 144, 146, 147, 148, 149, 152, 155, 156, 157, 158, 159, 160, 161, 162, 164, 165, 166, 167, 168, 169, 170, 172, 173, 174, 184, 192, 198, 199, 204, 205, 206, 207, 208, 209, 216)
    $apnicips = (1, 14, 27, 36, 39, 42, 43, 49, 58, 59, 60, 61, 101, 103, 106, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126, 133, 150, 153, 163, 171, 175, 180, 182, 183, 202, 203, 210, 211, 218, 219, 220, 221, 222, 223)
    $afrinicips = (41, 102, 105, 154, 196, 197)
    $lacnicips = (177, 179, 181, 186, 187, 189, 190, 191, 200, 201)
    $ripeips = (2, 5, 31, 37, 46, 62, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 109, 141, 145, 151, 176, 178, 185, 188, 193, 194, 195, 212, 213, 217)

    $octet = $ipaddress.split('{.}')

    if ($arinips -contains $octet[0])
    {
        return 'arin'
    }
    if ($apnicips -contains $octet[0])
    {
        return 'apnic'
    }
    if ($afrinicips -contains $octet[0])
    {
        return 'afrinic'
    }
    if ($lacnicips -contains $octet[0])
    {
        return 'lacnic'
    }
    if ($ripeips -contains $octet[0])
    {
        return 'ripe'
    }
}