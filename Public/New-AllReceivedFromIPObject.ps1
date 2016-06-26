#requires -Version 3
<#
        .Synopsis
        Short description
        .DESCRIPTION
        Long description
        .EXAMPLE
        Example of how to use this cmdlet
        .EXAMPLE
        Another example of how to use this cmdlet
#>
function New-AllReceivedFromIPObject
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true,
                ValueFromPipelineByPropertyName = $true,
        ParameterSetName = 'MessageObject')]
        [PSTypeName('PPRT.Message')]
        $MessageObject,

        [Parameter(Mandatory = $true,
                ValueFromPipelineByPropertyName = $true,
        ParameterSetName = 'EmailHeader')]
        $EmailHeader,

        [Parameter(Mandatory = $true,
        ValueFromPipelineByPropertyName = $true)]
        $SavePath,

        [Parameter(Mandatory = $false,
        ValueFromPipelineByPropertyName = $true)]
        [switch]$HeatMap
    )

    Begin
    {
        #regex is used for getting IPs from String
        $regex = '\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b'
        
        $polyline = @() #used for array of PolyLines
        $originalPolyline = @()
        $ReceivedFromIP = @()
        $ReturnObject = @()
    }
    Process
    {

        switch ($PSBoundParameters.Keys)
        {
            'MessageObject' 
            {
                $msg = $MessageObject.Headers 
            }
            'EmailHeader'   
            {
                if($null -ne $EmailHeader)
                {
                    $msg = $EmailHeader 
                }
                else
                {
                    Write-Warning -Message 'Please provide Email Headers'
                    break
                }
            }
        }

        foreach ($item in $MessageObject)
        {
            $ReceivedFromIP = (Parse-EmailHeader -InputFileName $item.Headers).From |
            Select-String -Pattern $regex -AllMatches |
            ForEach-Object -Process {
                $_.Matches 
            } |
            ForEach-Object -Process {
                $_.Value 
            }
            
        
            foreach ($ip in $ReceivedFromIP)
            {
                $IpLocation = ''
                $IpLocation = Invoke-RestMethod -Uri "http://freegeoip.net/xml/$($ip)"

                if (($IpLocation.Response.Latitude -ne 0) -or ($IpLocation.Response.Longitude -ne 0))
                {
                    if (![string]::IsNullOrWhiteSpace($IpLocation.Response.Latitude))
                    {
                        if (![string]::IsNullOrWhiteSpace($IpLocation.Response.Longitude))
                        {
                            $originalPolyline = "{lat: $($IpLocation.Response.Latitude), lng: $($IpLocation.Response.Longitude)}"
                            $polyline += $originalPolyline
                        }
                    }
                }
            }
        
            $props = @{
                marker          = "[$($polyline -join ',')]"
                subject         = $item.Subject
                SentFromAddress = $item.SenderEmailAddress
                SentFromType    = $item.SenderEmailType
                ReceivedTime    = $item.ReceivedTime
                EmailBody       = $item.Body
            }

            $tempAllIPObject = New-Object -TypeName PSObject -Property $props
            $AllIPObject += $tempAllIPObject

            $polyline = @()
        }

        $ReturnObject = $AllIPObject
    }
    End
    {
        return $ReturnObject
    }
}
