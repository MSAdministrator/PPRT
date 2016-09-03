function Check-ARIN (){

param (
        [parameter(Mandatory=$true,Position=1,HelpMessage="Please provide a IP address")]
        $ipaddress,

        [Parameter(Mandatory = $true)]
        [ValidateScript({ if (Test-Path $_){$true}else{ throw 'Please provide a valid path for LogPath' }})]
        $LogPath
            ) 
    <#
    .SYNOPSIS 
    Takes IP address as input and queries ARIN's WHOIS implementation for the IP addresses abuse contact email - RESTFul API 

    .DESCRIPTION
    Takes a IP Address and searches for ARIN's WHOIS abuse contact email for that IP based on registration data
    Returns this contact email address

    .PARAMETER ipaddress
    Specifices the specific IP address belonging to ARIN
   

    .EXAMPLE
    C:\PS> Check-ARIN 146.42.65.82   

    #>
    Begin
    {
        try
        {
            $TestingWebRequest = Invoke-WebRequest -Uri 'https://www.google.com'
        }
        catch
        {
            throw 'Unable to connect to google.com or Internet Explorer has not be used on this system.`
                   Please correct this before proceeding.'
            break
        }
    }
    Process
    {
        try
        {
            $ipdata = Invoke-RestMethod -Uri "http://whois.arin.net/rest/ip/$ipaddress" -UseBasicParsing
        }
        catch
        {
            Write-LogEntry -type ERROR `
                           -message 'Unable to Invoke-RestMethod against whois.arin.net' `
                           -CustomMessage $($Error[0] | Format-List -Property * -Force) `
                           -Folder $LogPath
        }

        try
        {
            $statusCode = Invoke-WebRequest $(($ipdata.net.orgRef.'#text')+'/pocs') | % {$_.StatusCode}
        }
        catch
        {
            Write-LogEntry -type ERROR `
                           -message 'Unable to Invoke-WebRequest to get Status code of site' `
                           -CustomMessage $($Error[0] | Format-List -Property * -Force) `
                           -Folder $LogPath
        }

        $orgdata = Invoke-RestMethod -Uri $(($ipdata.net.orgRef.'#text')+'/pocs')
        #$parentnetref = Invoke-RestMethod -Uri $(($ipdata.net.parentNetRef.'#text')+'/org/pocs')

       # $parentpocdata = Invoke-RestMethod -Uri  ("http://whois.arin.net/rest/poc/"+$($parentnetref.pocs.pocLinkRef | ?{$_.description -eq 'Abuse'}).handle)
       # $parentpocdata
        $orgDataObject = @()
        
        foreach ($item in $orgdata.pocs.pocLinkRef)
        {
            $pocdata = @()
            $pocdata = Invoke-RestMethod -Uri ("http://whois.arin.net/rest/poc/"+$($item.handle))#
            
            $props = @{
                Type = $item.Description
                Function = $item.Function
                Handle = $item.Handle
                URL = $item.'#text'
                POC = $pocdata.poc
            }
            
            $tempObject = New-Object -TypeName PSCustomObject -Property $props
            $orgDataObject += $tempObject
        }

        Add-ObjectDetail -InputObject $orgDataObject -TypeName PPRT.ARIN

    
        
   #     $pocdata = Invoke-RestMethod -Uri  ("http://whois.arin.net/rest/poc/"+$($orgdata.pocs.pocLinkRef)) #| ?{$_.description -eq 'Abuse'}).handle)
   #     $pocdata
   # If ($pocdata.poc.emails.InnerText -ne ""){return $pocdata.poc.emails.InnerText}
   # If ($parentpocdata.poc.emails.InnerText -ne ""){return $parentpocdata.poc.emails.InnerText}
    }
    End
    {
   #     return "NO ABUSE POC ON RECORD"
    }
}



