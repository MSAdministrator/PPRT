#requires -Version 2
function Get-IPaddress ()
{
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true,
        HelpMessage = 'Please provide a valid HOSTNAME')]
        $hostname,

        [Parameter(Mandatory = $true)]
        [ValidateScript({ if (Test-Path $_){$true}else{ throw 'Please provide a valid path for LogPath' }})]
        $LogPath
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
    try
    {
        $log = Write-LogEntry -type Info -message "Get-IPAddress: Attempting GetHostAddresses for $hostname" -Folder $LogPath
        $ipaddresses = [System.Net.Dns]::GetHostAddresses("$hostname").IPAddressToString
        $log = Write-LogEntry -type Info -message "Get-IPAddress: GetHostAddresses resolved $hostname to the following ipaddress(es)" -Folder $LogPath -CustomMessage $ipaddresses

        return $ipaddresses
    }
    catch
    {
        $log = Write-LogEntry -type Error -message "Get-IPAddress: GetHostAddresses could not resolve this host name: $hostname" -Folder $LogPath
    }

    try
    {
        $log = Write-LogEntry -type Info -message "Get-IPAddress: Attempting Test-Connection for $hostname" -Folder $LogPath
        $ipaddresses = (Test-Connection $hostname -Count 1).IPV4Address
        $log = Write-LogEntry -type Info -message "Get-IPAddress: Test-Connection resolved $hostname to the following ipaddress(es)" -Folder $LogPath -CustomMessage $ipaddresses
        
        return $ipaddresses
    }
    catch
    {
        $log = Write-LogEntry -type Error -message "Get-IPAddress: Test-Connection could not resolve this host name: $hostname" -Folder $LogPath
    }


    return $null
}
