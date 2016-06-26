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
function New-FirstReceivedFromIPObject
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
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

        $firstReceivedFromIP = @()
        $originalIpLocation = @()
        $originalmarker = @()
        $ReturnObject = @()
    }
    Process
    {
        switch ($PSBoundParameters.Keys)
        {
            'MessageObject' 
            {
                $msg = $MessageObject 
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

        foreach($item in $msg)
        {
            $firstReceivedFromIP = (Parse-EmailHeader -InputFileName $item.Header).From |
            `
            Select-String -Pattern $regex -AllMatches |
            `
            ForEach-Object -Process {
                $_.Matches 
            } |
            `
            ForEach-Object -Process {
                $_.Value 
            }
            
            #calling first received from header returned from parse-emailheader. Location is [0]
            $originalIpLocation = Invoke-RestMethod -Uri "http://freegeoip.net/xml/$($firstReceivedFromIP[0])"

            $tempStartingIPObject = @()

            #getting all first received from IP from headers and creating markers
            if (($originalIpLocation.Response.Latitude -ne 0) -or ($originalIpLocation.Response.Longitude -ne 0))
            {
                if (![string]::IsNullOrWhiteSpace($originalIpLocation.Response.Latitude))
                {
                    if (![string]::IsNullOrWhiteSpace($originalIpLocation.Response.Longitude))
                    {
                        #adding json markup data to object.  This will be passed to Get-PhishingGeoLocationStartingIps cmdlet
                        $props = @{
                            marker          = "`{'title': '$($item.subject -replace "'",' ')', `
                                'lat': '$($originalIpLocation.Response.Latitude)', `
                                'lng': '$($originalIpLocation.Response.Longitude)', `
                                'description': '<div><div></div><h1>$($item.Subject -replace "'",' ')</h1><div><p><b> `
                                Subject</b>: $($item.Subject -replace "'",' ')</p><p><b> `
                                Received Time</b>: $($item.ReceivedTime)</p><p><b> `
                                Sender Email Address</b>: $($item.SenderEmailAddress)</p><p><b> `
                                Sender Email Type</b>: $($item.SenderEmailType)</p><p><b> `
                                Phishing URL</b>: $($item.URL.RawPhishingLink)</p></div></div>' `
                            }"
                            subject         = $item.Subject
                            SentFromAddress = $item.SenderEmailAddress
                            SentFromType    = $item.SenderEmailType
                            ReceivedTime    = $item.ReceivedTime
                            EmailBody       = $item.Body
                        }

                        $tempStartingIPObject = New-Object -TypeName PSObject -Property $props
                    }
                }
            }

            $tempHeatMapObject = @()

            if ($HeatMap)
            {
                #getting heat map markers, even though they switch may not be called
                if (($originalIpLocation.Response.Latitude -ne 0) -or ($originalIpLocation.Response.Longitude -ne 0))
                {
                    if (![string]::IsNullOrWhiteSpace($originalIpLocation.Response.Latitude))
                    {
                        if (![string]::IsNullOrWhiteSpace($originalIpLocation.Response.Longitude))
                        {
                            $props = @{
                                marker          = "new google.maps.LatLng($($originalIpLocation.Response.Latitude), $($originalIpLocation.Response.Longitude))"
                                subject         = $item.Subject
                                SentFromAddress = $item.SenderEmailAddress
                                SentFromType    = $item.SenderEmailType
                                ReceivedTime    = $item.ReceivedTime
                                EmailBody       = $item.Body
                            }

                            $tempHeatMapObject = New-Object -TypeName PSObject -Property $props
                        }
                    }
                }
            }

            $props = @{
                FirstReceivedFromIP        = $tempStartingIPObject
                FirstReceivedFromIPHeatMap = $tempHeatMapObject
            }

            $TempObject = New-Object -TypeName PSObject -Property $props
            $ReturnObject += $TempObject
        }
    }
    End
    {
        return $ReturnObject
    }
}
