#requires -Version 3
function Get-PhishingGeoLocation 
{
    param (
        [parameter(Mandatory = $true,
        HelpMessage = 'Please provide a .msg file')]
        [ValidateNotNullOrEmpty()]
        [string[]]$email,
        [parameter(ParameterSetName = 'set2',
        HelpMessage = 'Please select either MapAllIPs or MapStartingIP')]
        [ValidateNotNullOrEmpty()]
        [switch]$MapAllIPs,
        [parameter(ParameterSetName = 'set1',
        HelpMessage = 'Please select either MapAllIPs or MapStartingIP')]
        [ValidateNotNullOrEmpty()]
        [switch]$MapStartingIP,
        [parameter(ParameterSetName = 'set1',
        HelpMessage = 'Please select either MapAllIPs or MapStartingIP')]
        [ValidateNotNullOrEmpty()]
        [switch]$HeatMap,
        [parameter(Mandatory = $true,
        HelpMessage = 'Please provide a folder path for ouputting generated maps')]
        [ValidateNotNullOrEmpty()]
        [string]$FolderPath

    ) 
    <#
            .SYNOPSIS 
            Map location of IPs from email headers of Phishing Emails received using Google Maps API

   
            .DESCRIPTION
            This function takes a .msg or .msg's as input and plots the email header IPs on local Google Map using the API. 
            This function takes the .msg, get's the email headers, get's all the IPs, finds the Latitude and Longtitude (if available)
            and builds a HTML webpage plotting those IPs

            You have several options for different types of maps:
                MapStartingIP - This options places markers for the first IP listed in the Received From headers of the email
                MapAllIPs - This option maps all "Received From" IPs from email headers and maps their total path
                HeatMap - This option is similar to MapStartingIp, but instead of markers it provides a Heat Map of those first received from IPs

            .PARAMETER message
            Specifices the message or messages you are wanting to plot on a map.  This messages need to be in a .msg format

            .PARAMETER MapAllIPs
            This switch is used when you want to map all IPs in a message header.  This option will 
            plot the path of the entire message header(s).

            .PARAMETER MapStartingIP
            This switch is used when you want to map the first "Received From:" IP.  This option will
            place a marker at it's geolocation.  The marker includes the following information:
            Subject
            Received Time
            Sender Email Address
            Phishing URL

            .PARAMETER HeatMap
            This switch can be used in conjunction with the MapStartingIP.  It will produce a Heat Map of the first "Received From:" IP. Coloration/Radius
            will increase based on the number of IPs

            .PARAMETER FolderPath
            This parameter is mandatory and is needed to to output the generated maps.

            .EXAMPLE
            C:\PS> Get-PhishingGeoLocation -message 'C:\users\username\PHISHING_MESSAGES\*.msg' -MapStartingIP -HeatMap -FolderPath C:\users\username\Desktop

            .EXAMPLE
            C:\PS> Get-PhishingGeoLocation -message 'C:\users\username\PHISHING_MESSAGES\*.msg' -MapStartingIP -FolderPath C:\users\username\Desktop

            .EXAMPLE
            C:\PS> Get-PhishingGeoLocation -message 'C:\users\username\PHISHING_MESSAGES\*.msg' -MapAllIPs -FolderPath C:\users\username\Desktop

            .NOTES
            Currently all Outputs work best in Internet Explorer but they may work in Google Chrome and FireFox.
    #>

    #setting arrays to be used later in the script
    $polyline = @() #used for array of PolyLines
    $StartingIPObject = @() #used for array of First Received IPs
    $HeatMapObject = @() #used for HeatMap Object
    $AllIPObject = @() #used for MapAllIPs object
    $i = 0 #used in conjunction with write-progress functionality
    $message = @() #used in conjunction with write-progress functionality

    #regex is used for getting IPs from String
    $regex = '\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b'
    
    #used to strip http/https/ftp/file/www links from emails
    $PhishingURLRegEx = '(?:(?:https?|ftp|file)://|www\.|ftp\.)(?:\([-A-Z0-9+&@#/%=~_|$?!:,.]*\)|[-A-Z0-9+&@#/%=~_|$?!:,.])*(?:\([-A-Z0-9+&@#/%=~_|$?!:,.]*\)|[A-Z0-9+&@#/%=~_|$])'

    #gathering total count of emails to process and adding them to $message array for further processing
    foreach ($item in $email)
    {
        Get-ChildItem $item | ForEach-Object -Process {
            $message += $_
        }
    }
    
    Write-Verbose -Message "Gathering Email Data from $(($message | Measure-Object).count) email messages"

    Get-ChildItem $message | ForEach-Object -Process {
        try 
        {
            Add-Type -AssemblyName 'Microsoft.Office.Interop.Outlook'
            $outlook = New-Object -ComObject outlook.application
        }
        catch
        {
    
            Write-Debug -Message 'Error: Please try and shutdown Outlook'
        }

        $msg = $outlook.CreateItemFromTemplate($_.FullName)

        #generating Progress Bar/Activity
        Write-Progress -Activity "Gathering Email Data from $($message.count) email messages" -Status "Processing $($msg.Subject)" -PercentComplete ($i/$(($message | Measure-Object).count)*100)

        #getting phishing URL from the current processing email message
        $phishingURL = ''

        $phishingURL = $msg | Select-Object -Property body |
                              Select-String -Pattern '(?:(?:https?|ftp|file)://|www\.|ftp\.)(?:\([-A-Z0-9+&@#/%=~_|$?!:,.]*\)|[-A-Z0-9+&@#/%=~_|$?!:,.])*(?:\([-A-Z0-9+&@#/%=~_|$?!:,.]*\)|[A-Z0-9+&@#/%=~_|$])' |
                              ForEach-Object -Process { $_.Matches } |
                              ForEach-Object -Process { $_.Value } 
        
        Write-Verbose -Message "PhisingURL: $($phishingURL)"

        $headers = ''

        #getting mapi property descriptor from microsoft.  This is needed to get the raw text email headers
        $headers = $msg.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x007D001E') 
     
        if ($MapStartingIP)
        {
            $firstReceivedFromIP = @()
            $firstReceivedFromIP = (Parse-EmailHeader -InputFileName $headers).From |
                Select-String -Pattern $regex -AllMatches | ForEach-Object -Process { $_.Matches } |
                    ForEach-Object -Process { $_.Value }
            
            #calling first received from header returned from parse-emailheader. Location is [0]
            $originalIpLocation = @()
            $originalIpLocation = Invoke-RestMethod -Uri "http://freegeoip.net/xml/$($firstReceivedFromIP[0])"
            
            #getting all first received from IP from headers and creating markers
            $originalmarker = @()
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
        }

        #getting data for MapAllIPs Google Maps API Polyline output
        if ($MapAllIPs)
        {
            $originalPolyline = @()
            $ReceivedFromIP = @()
            $ReceivedFromIP = (Parse-EmailHeader -InputFileName $headers).From | Select-String -Pattern $regex -AllMatches |
                ForEach-Object -Process { $_.Matches } |
                ForEach-Object -Process { $_.Value }

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
                subject         = $msg.Subject
                SentFromAddress = $msg.SenderEmailAddress
                SentFromType    = $msg.SenderEmailType
                ReceivedTime    = $msg.ReceivedTime
                EmailBody       = $msg.Body
            }

            $tempAllIPObject = New-Object -TypeName PSObject -Property $props
            $AllIPObject += $tempAllIPObject

            $polyline = @()
        }

        $i++
    }#end of foreach message

    if ($MapStartingIP)
    {
        Write-Verbose -Message "Starting IP Object Count: $(($StartingIPObject.marker).count)"
        Get-PhishingGeoLocationStartingIPs -StartingIPData $StartingIPObject -FolderPath $FolderPath
    }

    if ($HeatMap)
    {
        Write-Verbose -Message "Heat Map Object Count: $(($HeatMapObject.marker).count)"
        Get-PhishingGeoLocationHeatMap -HeatMapData $HeatMapObject -FolderPath $FolderPath
    }
    
    if ($MapAllIPs)
    {
        Write-Verbose -Message "All IP Object Count: $(($AllIPObject.marker).count)"
        Get-PhishingGeoLocationAllIPs -AllIPData $AllIPObject -FolderPath $FolderPath
    }

    #closing outlook process if not done already
    Start-Sleep -Seconds 3
    Get-Process -Name Outlook | Stop-Process
}
