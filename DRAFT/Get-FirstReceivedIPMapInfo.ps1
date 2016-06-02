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
function Get-FirstReceivedIPMapInfo
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $headers,

        # Param2 help description
        [int]
        $Param2
    )

    Begin
    {
        $firstReceivedFromIP = @()
        $originalIpLocation = @()
        $originalmarker = @()
        $tempHeatMapMarkers = @()
    }
    Process
    {
        $firstReceivedFromIP = (Parse-EmailHeader -InputFileName $headers).From |
            Select-String -Pattern $regex -AllMatches | ForEach-Object -Process { $_.Matches } |
                ForEach-Object -Process { $_.Value }

        #calling first received from header returned from parse-emailheader. Location is [0]
        $originalIpLocation = Invoke-RestMethod -Uri "http://freegeoip.net/xml/$($firstReceivedFromIP[0])"
            
        #getting all first received from IP from headers and creating markers
        if (($originalIpLocation.Response.Latitude -ne 0) -or ($originalIpLocation.Response.Longitude -ne 0))
        {
            if (![string]::IsNullOrWhiteSpace($originalIpLocation.Response.Latitude))
            {
                if (![string]::IsNullOrWhiteSpace($originalIpLocation.Response.Longitude))
                {
                    #adding json markup data to object.  This will be passed to Get-PhishingGeoLocationStartingIps cmdlet
                    $props = @{
                        marker          = "`{'title': '$($msg.subject -replace "'",' ')', 'lat': '$($originalIpLocation.Response.Latitude)', 'lng': '$($originalIpLocation.Response.Longitude)', 'description': '<div><div></div><h1>$($msg.Subject -replace "'",' ')</h1><div><p><b>Subject</b>: $($msg.Subject -replace "'",' ')</p><p><b>Received Time</b>: $($msg.ReceivedTime)</p><p><b>Sender Email Address</b>: $($msg.SenderEmailAddress)</p><p><b>Sender Email Type</b>: $($msg.SenderEmailType)</p><p><b>Phishing URL</b>: $($phishingURL)</p></div></div>' }"
                        subject         = $msg.Subject
                        SentFromAddress = $msg.SenderEmailAddress
                        SentFromType    = $msg.SenderEmailType
                        ReceivedTime    = $msg.ReceivedTime
                        EmailBody       = $msg.Body
                    }

                    $tempStartingIPObject = New-Object -TypeName PSObject -Property $props
                    $StartingIPObject += $tempStartingIPObject
                }
            }
        }

        #getting heat map markers, even though they switch may not be called
        if (($originalIpLocation.Response.Latitude -ne 0) -or ($originalIpLocation.Response.Longitude -ne 0))
        {
            if (![string]::IsNullOrWhiteSpace($originalIpLocation.Response.Latitude))
            {
                if (![string]::IsNullOrWhiteSpace($originalIpLocation.Response.Longitude))
                {
                    $props = @{
                        marker          = "new google.maps.LatLng($($originalIpLocation.Response.Latitude), $($originalIpLocation.Response.Longitude))"
                        subject         = $msg.Subject
                        SentFromAddress = $msg.SenderEmailAddress
                        SentFromType    = $msg.SenderEmailType
                        ReceivedTime    = $msg.ReceivedTime
                        EmailBody       = $msg.Body
                    }

                    $tempHeatMapObject = New-Object -TypeName PSObject -Property $props
                    $HeatMapObject += $tempHeatMapObject
                }
            }
        }
    }
    End
    {
        $props = @{
            StartingIPObject = $StartingIPObject
            HeatMapObject = $HeatMapObject
        }
        
        $ReturnObject = New-Object -TypeName PSObject -Property $props

        return $ReturnObject
    }
}





            $firstReceivedFromIP = (Parse-EmailHeader -InputFileName $headers).From |
                Select-String -Pattern $regex -AllMatches | ForEach-Object -Process { $_.Matches } |
                    ForEach-Object -Process { $_.Value }
            
            #calling first received from header returned from parse-emailheader. Location is [0]
      
            $originalIpLocation = Invoke-RestMethod -Uri "http://freegeoip.net/xml/$($firstReceivedFromIP[0])"
            
            #getting all first received from IP from headers and creating markers
            if (($originalIpLocation.Response.Latitude -ne 0) -or ($originalIpLocation.Response.Longitude -ne 0))
            {
                if (![string]::IsNullOrWhiteSpace($originalIpLocation.Response.Latitude))
                {
                    if (![string]::IsNullOrWhiteSpace($originalIpLocation.Response.Longitude))
                    {
                        #adding json markup data to object.  This will be passed to Get-PhishingGeoLocationStartingIps cmdlet
                       $props = @{
                            marker          = "`{'title': '$($msg.subject -replace "'",' ')', 'lat': '$($originalIpLocation.Response.Latitude)', 'lng': '$($originalIpLocation.Response.Longitude)', 'description': '<div><div></div><h1>$($msg.Subject -replace "'",' ')</h1><div><p><b>Subject</b>: $($msg.Subject -replace "'",' ')</p><p><b>Received Time</b>: $($msg.ReceivedTime)</p><p><b>Sender Email Address</b>: $($msg.SenderEmailAddress)</p><p><b>Sender Email Type</b>: $($msg.SenderEmailType)</p><p><b>Phishing URL</b>: $($phishingURL)</p></div></div>' }"
                            subject         = $msg.Subject
                            SentFromAddress = $msg.SenderEmailAddress
                            SentFromType    = $msg.SenderEmailType
                            ReceivedTime    = $msg.ReceivedTime
                            EmailBody       = $msg.Body
                        }

                        $tempStartingIPObject = New-Object -TypeName PSObject -Property $props
                        $StartingIPObject += $tempStartingIPObject
                    }
                }
            }

            #getting heat map markers, even though they switch may not be called
            $tempHeatMapMarkers = @()
            if (($originalIpLocation.Response.Latitude -ne 0) -or ($originalIpLocation.Response.Longitude -ne 0))
            {
                if (![string]::IsNullOrWhiteSpace($originalIpLocation.Response.Latitude))
                {
                    if (![string]::IsNullOrWhiteSpace($originalIpLocation.Response.Longitude))
                    {
                     $props = @{
                            marker          = "new google.maps.LatLng($($originalIpLocation.Response.Latitude), $($originalIpLocation.Response.Longitude))"
                            subject         = $msg.Subject
                            SentFromAddress = $msg.SenderEmailAddress
                            SentFromType    = $msg.SenderEmailType
                            ReceivedTime    = $msg.ReceivedTime
                            EmailBody       = $msg.Body
                        }

                        $tempHeatMapObject = New-Object -TypeName PSObject -Property $props
                        $HeatMapObject += $tempHeatMapObject
                    }
                }
            }