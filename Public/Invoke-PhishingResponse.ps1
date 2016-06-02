<#
    .SYNOPSIS 
    Takes a .msg file, find a phishing link, does reverse DNS for the IP, and queries whois Databases for abuse contact information

    .DESCRIPTION
    Takes a .MSG file and searches for a link based on a regex pattern
    Takes that link, parses it to find the root DNS name
    Takes the DNS name and finds the IP by doing a reverse DNS lookup
    Takes the IP of the server and parses it for the first octet
    Takes the first octet and finds which whois should be used
    Once it has the whois, it queries their API or scraps their website for their abuse contact information
    Once it has the abuse contact info, it sends them an email from abuse email account with the original attachment - asking them to remove the website
    Sends an email to spam@access.ironport.com
    Sends an email to the Google Anti-Phishing Group anti-phishing-email-reply-discuss@googlegroups.com
    Logs this in the running log file

    .PARAMETER messagetoparse
    Specifices the specific .MSG that someone wants to parse 

    .PARAMETER logpath
    Sets the path to our log file

    .PARAMETER From
    This parameter is used to define who is sending these notificaitons.
    Currently, you must put an email address that you want to "Send on Behalf of".

    .EXAMPLE
    C:\PS> Send-PhishingNotification -meesagetoparse 'C:\Users\UserName\Desktop\PHISING_EMAILS\Dear Email User.msg' -logpath C:\users\username\desktop -From 'abuse@emailaddress.com'

#>

#requires -Version 2
function Invoke-PhishingResponse
{
	[CmdletBinding()]
    param (
        [parameter(Mandatory = $true,
        HelpMessage = 'Please provide a .MSG file.')]
        $MailMessage,

        [parameter(Mandatory = $true,
        HelpMessage = "Please provide a 'Send On Behalf of' email address")]
        $From,

        [parameter(ParameterSetName='Set3',
        HelpMessage = "Please include the VirusTotal switch to scan files against VT API.")]
        [switch]$VirusTotal,

        [parameter(ParameterSetName='Set3',
        HelpMessage = "Please provide your Virus Total API Key")]
        $VTAPIKey,

        [parameter(HelpMessage = "Provide this switch if you want to send additional notifications")]
        [switch]$AdditionalNotifications,

        [parameter(ParameterSetName = 'Set1',
        HelpMessage = 'Please select either AllReceivedFromIPMap or FirstReceivedFromIPMap')]
        [ValidateNotNullOrEmpty()]
        [switch]$AllReceivedFromIPMap,

        [parameter(ParameterSetName = 'Set2',
        HelpMessage = 'Please select either AllReceivedFromIPMap or FirstReceivedFromIPMap')]
        [ValidateNotNullOrEmpty()]
        [switch]$FirstReceivedFromIPMap,

        [parameter(ParameterSetName = 'Set2',
        HelpMessage = 'Please select either AllReceivedFromIPMap or FirstReceivedFromIPMap')]
        [ValidateNotNullOrEmpty()]
        [switch]$FirstReceivedFromIPHeatMap
    ) 

    Begin
    {
        if ((Get-Item $MailMessage) -is [System.IO.DirectoryInfo])
        {
            $LogLocation = $MailMessage
        }
        else
        {
            $LogLocation = Split-Path -Parent $MailMessage
        }
        
        

        $ipaddress = @()
        $regexipv6 = '(([0-9a-fA-F]{1,4}:){7,7}[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,7}:|([0-9a-fA-F]{1,4}:){1,6}:[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,5}(:[0-9a-fA-F]{1,4}){1,2}|([0-9a-fA-F]{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}|([0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}|([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4}){1,5}|[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4}){1,6})|:((:[0-9a-fA-F]{1,4}){1,7}|:)|fe80:(:[0-9a-fA-F]{0,4}){0,4}%[0-9a-zA-Z]{1,}|::(ffff(:0{1,4}){0,1}:){0,1}((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])|([0-9a-fA-F]{1,4}:){1,4}:((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]))'
        $regexipv4 = '\b((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3} (25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b'
        $URLRegEx = '(?:(?:https?|ftp|file)://|www\.|ftp\.)(?:\([-A-Z0-9+&@#/%=~_|$?!:,.]*\)|[-A-Z0-9+&@#/%=~_|$?!:,.])*(?:\([-A-Z0-9+&@#/%=~_|$?!:,.]*\)|[A-Z0-9+&@#/%=~_|$])'
        $shorturl = $false
        $URLObject = @()
        $message = @()
        $PhishingURL = @()
        $PhisingAttachment = @()

        #gathering total count of emails to process and adding them to $message array for further processing
        foreach ($item in $MailMessage)
        {
            Get-ChildItem $item | ForEach-Object -Process {
                $message += $_
            }
        }
    
        Write-LogEntry -type Info -message "Gathering message count" -Folder $LogLocation -CustomMessage "Total Message Count is $(($message | Measure-Object).Count)"

        try 
        {
            Add-Type -AssemblyName 'Microsoft.Office.Interop.Outlook'
            $outlook = New-Object -ComObject outlook.application
        }
        catch
        {
            Write-LogEntry -type Error -message "Can not load Microsoft.Office.Interop.Outlook" -Folder $LogLocation
            Write-LogEntry -type Error -message "Exiting...." -Folder $LogLocation
            Exit
        }
    }
    Process
    {
        $MainObject = @()

        Get-ChildItem $MailMessage | ForEach-Object -Process {

            $msgFn = $_.FullName

            # Skip non-.msg files
            if ($msgFn -like "*.msg") {

                Write-Host "LogLocation: $LogLocation"
                Write-LogEntry -type Info -message "Processing Phishing Email - $($_.Subject)" -Folder $LogLocation
                
                $MessageObject = @()
                $AttachmentObject = @()
                # Work out file names

                $msg = $outlook.CreateItemFromTemplate($msgFn)

                Write-LogEntry -type Info -message "Building Phishing Email Object" -Folder $LogLocation

                $MessageProperties = @{
                    FullName = $msgFn
                    Subject = $msg.Subject
                    Body = $msg.Body
                    HTMLBody = $msg.HTMLBody
                    BCC = $msg.BCC
                    CC = $msg.CC
                    ReceivedOnBehalfOfEntryID = $msg.ReceivedOnBehalfOfEntryID
                    ReceivedOnBehalfOfName = $msg.ReceivedOnBehalfOfName
                    ReceivedTime = $msg.ReceivedTime
                    Receipents = $msg.Recipients
                    ReplyRecipientsName = $msg.ReplyRecipientNames
                    SenderName = $msg.SenderName
                    SentOnDate = $msg.SentOn
                    SentOnBehalfOfName = $msg.SentOnBehalfOfName
                    SentTo = $msg.To
                    SenderEmailAddress = $msg.SenderEmailAddress
                    SenderEmailType = $msg.SenderEmailType
                    SendUsingAccount = $msg.SendUsingAccount
                    Attachments = ''
                    URL = ''
                    Headers = ''
                }

                $MessageObject = New-Object -TypeName PSObject -Property $MessageProperties

                #generating Progress Bar/Activity
                Write-Progress -Activity "Gathering Email Data from $($MailMessage.count) email messages" -Status "Processing $($msg.Subject)" -PercentComplete ($i/$(($MailMessage | Measure-Object).count)*100)

                if ($msg.Attachments)
                {
                    
                    $PhisingAttachment = @()

                    $msg.Attachments | ForEach-Object -Process {

                        Write-LogEntry -type Info -message "Phishing Email Contains an Attachment - Processing - $($_.FileName)" -Folder $LogLocation
                        $attFn = ''

                        $AttachmentExists = $false

                        # Work out attachment file name
                        $attFn = $msgFn -replace '\.msg$', " - Attachment - $($_.FileName)"
                        Write-Verbose "Attachment File Name: $attFn"

                        # Do not try to overwrite existing files
                        if (Test-Path -literalPath $attFn) {
                            $AttachmentExists = $true
                            Write-LogEntry -type Info -message "Phishing Email Attachment already exists - $($attFn)" -Folder $LogLocation
                            Write-LogEntry -type Info -message "Skipping...." -Folder $LogLocation
                            break
                        }

                        
                        if (!(Test-Path -Path "$LogLocation\Extracted_Attachments"))
                        {
                            try
                            {
                                New-Item "$LogLocation\Extracted_Attachments" -ItemType Directory -Force 
                                Write-LogEntry -type Info -message "Created new folder to save Phishing Email Attachments - $("$LogLocation\Extracted_Attachments")" -Folder $LogLocation
                            }
                            catch
                            {
                                Write-LogEntry -type Error -message "Can not create folder for extracted attachments - $("$LogLocation\Extracted_Attachments")" -Folder $LogLocation
                                Write-LogEntry -type Error -message "Skipping Attachment Extraction....." -Folder $LogLocation
                                break
                            }
                        }

                        # Save attachment
                        Write-LogEntry -type Info -message "Saving Phishing Email Attachment - $("$LogLocation\Extracted_Attachments\$($_.FileName)")" -Folder $LogLocation
                        [string]$SavePath = $("$LogLocation\Extracted_Attachments\$($_.FileName)")
                
                        $_.SaveAsFile($SavePath)

                        $AttachmentHash = Get-FileHash -Path $SavePath
                        Write-LogEntry -type Info -message "Phishing Email Attachment Hash - $($AttachmentHash)" -Folder $LogLocation

                        if ($VirusTotal)
                        {
                            Write-LogEntry -type Info -message "Checking Attachment against VirusTotal....." -Folder $LogLocation
                            $VirusTotalResult = Invoke-VirusTotal -AttachmentHash $AttachmentHash -VTAPIKey $VTAPIKey
                            Write-LogEntry -type Info -message "Processed Attachment against VirusTotal - Results - $($VirusTotalResult)" -Folder $LogLocation
                        }
                        $Props = @{
                            Message = $msg
                            OriginalAttachmentName = $($_.FileName)
                            NewAttachmentName = $($attFn)
                            AttachmentSavePath = $SavePath
                            AttachmentHash = $AttachmentHash
                            VirusTotalResults = $VirusTotalResult
                        }

                        $TempObject = New-Object -TypeName PSObject -Property $Props
                        Write-LogEntry -type Info -message "Phishing Email Attachment Object Created...." -Folder $LogLocation

                        $AttachmentObject += $TempObject
                    }
                }
                
                Write-LogEntry -type Info -message "Adding Attachment Object to Object...." -Folder $LogLocation
                #Adding AttachmentObject to custom MessageObject
                $MessageObject.Attachments = $AttachmentObject

                $headers = ''

                Write-LogEntry -type Info -message "Processing Email Headers...." -Folder $LogLocation

                #getting mapi property descriptor from microsoft.  This is needed to get the raw text email headers
                $headers = $msg.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x007D001E') 
                
                #Adding Message Headers to custom MessageObject
                $MessageObject.Headers = $headers


                Write-LogEntry -type Info -message "Getting URL(s) from Phishing Email...." -Folder $LogLocation
                #Getting the URL from message
                $PhishingURL = $msg |
                Select-Object -Property body |
                Select-String -Pattern $URLRegEx |
                ForEach-Object -Process {
                    $_.Matches
                } |
                ForEach-Object -Process {
                    $_.Value
                }

                $PhishingNotificationObject = @()

                foreach ($url in $PhishingURL)
                {
                    Write-LogEntry -type Info -message "Starting Phishing Notification Processing on URL - $($url)" -Folder $LogLocation
                    $URLObject = @()
                    $URLStatus = $false
                    $AbuseNotification = ''
                    $IronPortNotification = ''
                    $AntiPhishingNotification = ''

                    $URLAlive = (Invoke-WebRequest -Uri $url | select -Property StatusCode).StatusCode | Out-Null

                    if ($URLAlive -eq 200)
                    {
                        Write-LogEntry -type Info -message "Phishing URL Alive - $($url)" -Folder $LogLocation

                        $URLStatus = $true

                        Import-Clixml -Path "Z:\_Box\_GitHub\PhishReporter\Private\shorturls.xml" | ForEach-Object {
                            if ($url -like '$_')
                            {
                                Write-LogEntry -type Info -message "Phishing URL is a shortened URL - $($url)" -Folder $LogLocation
                                $shorturl = $true
                                $URLObject = Get-LongUrl $url -logpath $LogLocation
                            }
                        }

                        if ($shorturl -eq $false)
                        {
                            Write-LogEntry -type Info -message "Phishing URL is a long URL - $($url)" -Folder $LogLocation
                            $URLObject = Get-ParsedURL -url $url -logpath $LogLocation
                        }

                        Write-LogEntry -type Info -message "Gathering IP information for URL...." -Folder $LogLocation
                        [array]$ipaddress = Get-IPaddress -hostname $($URLObject).URLAuthority

                        #for each ipaddress returned from above statement
                        for ($ip = 0;$ip -lt $ipaddress.count;$ip++)
                        {
                            Write-LogEntry -type Info -message "Processing IP Address - $($ip)" -Folder $LogLocation
                            #based on the ipaddress we are going to get which WHOIS/RDAP to use
                            $whoisdb = Get-WhichWHOIS $ipaddress[$ip]
    
                            Write-LogEntry -type Info -message "IP Address belongs to - $($whoisdb)" -Folder $LogLocation
                            #based on info from Get-WhichWHOIS we will then begin those specific API calls
                            switch ($whoisdb){
                                'arin' 
                                {
                                    [array]$abusecontact = Check-ARIN $ipaddress[$ip]
                                    Write-LogEntry -type Info -message "Abuse Contact Information for Arin - $($abusecontact)" -Folder $LogLocation
                                }
                                'ripe' 
                                {
                                    [array]$abusecontact = Check-RIPE $ipaddress[$ip]
                                    Write-LogEntry -type Info -message "Abuse Contact Information for Ripe - $($abusecontact)" -Folder $LogLocation
                                }
                                'apnic' 
                                {
                                    $abusecontact = Check-APNIC $ipaddress[$ip]
                                    Write-LogEntry -type Info -message "Abuse Contact Information for Apnic - $($abusecontact)" -Folder $LogLocation
                                }
                                'lacnic' 
                                {
                                    [array]$abusecontact = Check-LACNIC $ipaddress[$ip]
                                    Write-LogEntry -type Info -message "Abuse Contact Information for Lacnic - $($abusecontact)" -Folder $LogLocation
                                }
                                'afrnic' 
                                {
                                    $abusecontact = 'NOCONTACT'
                                    Write-LogEntry -type Error -message "NO Abuse Contact Information Found!!!" -Folder $LogLocation
                                }
                                $null 
                                {

                                }
                            }
                        }

                        #as long as the abusecontact does not equal 'NOCONTACT', send email to that abuse contact
                        for ($a = 0;$a -lt $abusecontact.count;$a++)
                        {
                            if ($abusecontact[$a] -ne 'NOCONTACT') 
                            {
                                Write-LogEntry -type Info -message "Sending Notification to Abuse Contact" -Folder $LogLocation
                              #  $AbuseNotification = Send-ToAbuseContact -originallink $URLObject.OriginalURL -abusecontact $abusecontact[$a] -messagetoattach $MailMessage -From $From -LogLocation $LogLocation
                            }
                        }

                        if ($AdditionalNotifications)
                        {
                            Write-LogEntry -type Info -message "Sending Additional Notifications" -Folder $LogLocation
                            #additionally, send to IronPort and Anti-Phishing Working Group email distribution list
	                       # $IronPortNotification = Send-ToIronPort -originallink $url.OriginalURL -messagetoattach $MailMessage -From $From -LogLocation $LogLocation
	
	                     #   $AntiPhishingNotification = Send-ToAntiPhishingGroup -trimmedlink $url.URLAuthority -From $From -LogLocation $LogLocation
                        }
                    }

                    $URLProps = @{
                        RawPhishingLink = $url
                        PhishingLinkStatus = $URLStatus
                        ShortenedURL = $shorturl
                        URLObject = $URLObject
                        URLIPAddress = $ipaddress
                        WHOISInfo = @{
                                WhichWHOIS = $whoisdb
                                AbuseContact = $abusecontact
                            }
                        AbuseNotificationStatus = $AbuseNotification
                        AdditionalNotification = @{
                            IronPortNotification = $IronPortNotification
                            AntiPhishingNotification = $AntiPhishingNotification
                            }
                    }

                    $TempObject = New-Object -TypeName PSObject -Property $URLProps
                    Write-LogEntry -type Info -message "Phishing Notification Object Created" -Folder $LogLocation
                    $PhishingNotificationObject += $TempObject
                }

                #Adding PhishingURL(s) to custom MessageObject
                $MessageObject.URL = $PhishingNotificationObject
                Write-LogEntry -type Info -message "Adding Phishing Object to the Main Object" -Folder $LogLocation
                $MainObject += $MessageObject
            }#End of if Message is .msg
            {
                Write-LogEntry -type Error -message "ITEM NOT A MSG FILE - $msgFn" -Folder $LogLocation
                Write-LogEntry -type Error -message "Skipping...." -Folder $LogLocation
                break
            }

        }#end of processing all messages

        Write-LogEntry -type Info -message "All Phishing Emails Processed" -Folder $LogLocation

        if ($FirstReceivedFromIPMap)
        {
            if ($FirstReceivedFromIPHeatMap)
            {
                $FirstReceivedFromIPObject = Create-FirstReceivedFromIPObject -MessageHeaders $headers -SavePath $LogLocation -MessageObject $MainObject -HeatMap
            }
            else
            {
                $FirstReceivedFromIPObject = Create-FirstReceivedFromIPObject -MessageHeaders $headers -SavePath $LogLocation -MessageObject $MainObject
            }     
        }
                

        #getting data for MapAllIPs Google Maps API Polyline output
        if ($AllReceivedFromIPMap)
        {
            $AllReceivedFromIPObject = Create-AllReceivedFromIPObject -MessageObject $MainObject -SavePath $LogLocation
        }

    }
    End
    {
        #stop outlook process if still open from send emails using Outlook.Application COM Object
        Start-Sleep -Seconds 3
        Get-Process -Name Outlook | Stop-Process

        return $MainObject
    }
}
