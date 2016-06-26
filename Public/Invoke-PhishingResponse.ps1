#requires -Version 2
function Invoke-PhishingResponse
{
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

            .EXAMPLE 1
            C:\PS> Invoke-PhishingResponse -Message $MessageObject `
            -LogPath C:\PHISHING_EMAILS `
            -From 'abuse@emailaddress.com' `
            -ExtractAttachments `
            -SaveLocation C:\PHISHING_EMAILS\EXTRACTED_ATTACHMENTS

            .EXAMPLE 2
            C:\PS> Invoke-PhishingResponse -Message $MessageObject `
            -LogPath C:\PHISHING_EMAILS `
            -From 'abuse@emailaddress.com' `
            -ExtractAttachments `
            -SaveLocation C:\PHISHING_EMAILS\EXTRACTED_ATTACHMENTS `
            -SMTPServer smtp.office365.com `
            -Credential $Cred

    #>

    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true,
                HelpMessage = 'Please provide a .MSG file.',
                ValueFromPipelineByPropertyName = $true,
        ValueFromPipeline = $true)]
        [PSTypeName('PPRT.Message')]
        $Message,

        [Parameter(Mandatory = $true)]
        $LogPath,

        [Parameter(Mandatory = $true,
        ParameterSetName = 'VT')] 
        [switch]$ExtractAttachments,

        [Parameter(Mandatory = $true,
        ParameterSetName = 'VT')]
        [Parameter(Mandatory = $true,
        ParameterSetName = 'Map1')]
        [string]$SaveLocation,

        [parameter(Mandatory = $true,
                HelpMessage = 'Please provide a From email address',
        ParameterSetName = 'Email')]
        [string]$From,

        [parameter(Mandatory = $false,
                HelpMessage = 'Please provide a message Subject. Default uses the Phising Email subject.',
        ParameterSetName = 'Email')]
        $Subject,


        [Parameter(Mandatory = $true,
                ValueFromPipelineByPropertyName = $true,
        ParameterSetName = 'Email')]
        [Alias('PSEmailServer')]
        [string]$SMTPServer = $PSEmailServer,

        [Parameter(Mandatory = $false,
                ValueFromPipelineByPropertyName = $true,
        ParameterSetName = 'Email')]
        [int]$SMTPPort = '25',

        [Parameter(Mandatory = $false,
                ValueFromPipelineByPropertyName = $true,
        ParameterSetName = 'Email')]
        [switch]$UseSSL,

        [parameter(Mandatory = $false,
                HelpMessage = 'Please provide a message Body. Default uses just the phishing URL.',
        ParameterSetName = 'Email')]
        $Body,

        [parameter(Mandatory = $false,
                HelpMessage = 'Please indicate if you want the message body to be rendered as HTML.',
        ParameterSetName = 'Email')]
        [switch]$BodyAsHTML,

        [parameter(Mandatory = $true,
                HelpMessage = 'Please provide a message Body. Default uses just the phishing URL.',
        ParameterSetName = 'Email')]
        [System.Management.Automation.PSCredential]$Credential,

        [Parameter(ParameterSetName = 'VT')]
        [parameter(HelpMessage = 'Please provide your Virus Total API Key')]
        $VTAPIKey,

        [parameter(HelpMessage = 'Provide this switch if you want to send additional notifications')]
        [switch]$AdditionalNotifications,

        [parameter(HelpMessage = 'Please select either AllReceivedFromIPMap or FirstReceivedFromIPMap')]
        [ValidateNotNullOrEmpty()]
        [switch]$AllReceivedFromIPMap,

        [parameter(HelpMessage = 'Please select either AllReceivedFromIPMap or FirstReceivedFromIPMap')]
        [ValidateNotNullOrEmpty()]
        [Parameter(ParameterSetName = 'Map1')]
        [switch]$FirstReceivedFromIPMap,

        [parameter(HelpMessage = 'Please select either AllReceivedFromIPMap or FirstReceivedFromIPMap')]
        [ValidateNotNullOrEmpty()]
        [Parameter(ParameterSetName = 'Map1')]
        [switch]$FirstReceivedFromIPHeatMap
    ) 

    Begin
    {       
        $AttachmentObject = @()
        $VirusTotalResult = @()
        $URLObject = @()
        $NotificationObj = @()

        $MainObject = @()

        $PhishingResponseObject = @{}

    }
    Process
    {

        foreach ($msg in $Message)
        {
            if ($ExtractAttachments)
            {
                if ($msg.Attachments)
                {
                    #Call Extract-MessageAttachment
                    $AttachmentObject = Export-MessageAttachment -MessageObject $msg -LogPath $LogPath -FullDetails -SavePath $SaveLocation

                    $log = Write-LogEntry -type Info -message 'Invoke-PhishingResponse: Attachment Extracted' -Folder $LogPath
  
                    if ($VTAPIKey)
                    {           
                        $log = Write-LogEntry -type Info -message 'Invoke-PhishingResponse: Calling Invoke-VTAttachment' -Folder $LogPath          
                        $VirusTotalResult = Invoke-VTAttachment -AttachmentHash $AttachmentObject -VTAPIKey $VTAPIKey
                    }
                }
            }


            $URLObject = @()

            $log = Write-LogEntry -type Info -message "Invoke-PhishingResponse: Getting URL From $($msg.FullName)" -Folder $LogPath

            $URLObject = Get-URLFromMessage -MessageObject $msg -LogPath $LogPath

            $AbuseContactObject = @()

            $log = Write-LogEntry -type Info -message "Invoke-PhishingResponse: Trying to identify Abuse Contact for $($msg.FullName)" -Folder $LogPath

            $AbuseContactObject = New-PPRTAbuseContactObject -URLObject $URLObject -LogPath $LogPath
            
            if ($AbuseContactObject.AbuseContact -notmatch 'NO POC FOR *')
            {
                if ($AbuseContactObject -eq $null)
                {
                    $log = Write-LogEntry -type Info -message 'Invoke-PhishingResponse: New-PPRTAbuseContactObject did not return a value' -Folder $LogPath
                    continue
                }

                $Obj = @{}

                if ($null -ne $AbuseContactObject)
                {
                    $SendTo = @()
                    
                    foreach ($item in $AbuseContactObject.AbuseContact)
                    {
                        $SendTo += $($item)
                    }

                    $Obj.To = $($SendTo -join ',')
                }
                else
                {
                    $log = Write-LogEntry -type Error -message 'Invoke-PhishingResponse: ABUSE CONTACT IS NULL' -Folder $LogPath -CustomMessage 'Break!'
                    continue
                }

                if (!$Subject)
                {
                    $Obj.Subject = $msg.Subject
                }

                if (!$Body)
                {
                    $Obj.Body = $URLObject.URL
                }

                switch ($psboundparameters.keys) 
                {
                    'From'         
                    {
                        $Obj.From         = $From
                    }
                    'HTMLBody'     
                    {
                        $Obj.BodyAsHTML   = $BodyAsHTML
                    }
                    'BCC'          
                    {
                        $Obj.BCC          = $BCC
                    }
                    'CC'           
                    {
                        $Obj.CC           = $CC
                    }
                    'Subject'      
                    {
                        $Obj.Subject      = $Subject
                    }
                    'Priority'     
                    {
                        $Obj.Priority     = $Priority
                    }
                    'UseSSL'       
                    {
                        $Obj.UseSSL       = $UseSSL
                    }
                    'Encoding'     
                    {
                        $Obj.Encoding     = $Encoding
                    }
                    'Credential'   
                    {
                        $Obj.Credential   = $Credential
                    }
                    'SMTPServer'   
                    {
                        $Obj.SMTPServer   = $SMTPServer
                    }
                    'SMTPPort'     
                    {
                        $Obj.SMTPPort     = $SMTPPort
                    }
                    'URL'          
                    {
                        $Obj.URL          = $AbuseContactObject.URL 
                    }
                }

                $NotificationObject = @()

                $log = Write-LogEntry -type Info -message "Invoke-PhishingResponse: Attempting to send notifications - $($Obj)" -Folder $LogPath

                $NotificationObject = Send-MailMessage @Obj

                $PhishingResponseObject.Notification = $NotificationObj 
            }

            $props = @{
                MSG          = $msg
                URL          = $URLObject
                Abuse        = $AbuseContactObject
                Notification = $Obj
                Attachment   = $AttachmentObject
                VirusTotal   = $VirusTotalResult
            }

            $TempMainObject = New-Object -TypeName PSObject -Property $props
            $MainObject += $TempMainObject
        
            $log = Write-LogEntry -type Info -message "Invoke-PhishingResponse: Message Successfully Processed - $($msg.FullName)" -Folder $LogPath  
        }

        $log = Write-LogEntry -type Info -message 'All Phishing Emails Processed' -Folder $LogLocation

        if ($FirstReceivedFromIPMap)
        {
            if ($FirstReceivedFromIPHeatMap)
            {
                $log = Write-LogEntry -type Info -message 'Invoke-PhishingResponse: Creating New First Received From IP Heat Map Object' -Folder $LogLocation
                $FirstReceivedFromIPObject = New-FirstReceivedFromIPObject -MessageObject $Message -SavePath $SaveLocation -HeatMap
            }
            else
            {
                $log = Write-LogEntry -type Info -message 'Invoke-PhishingResponse: Creating New First Received From IP Map Object' -Folder $LogLocation
                $FirstReceivedFromIPObject = New-FirstReceivedFromIPObject -MessageObject $Message -SavePath $SaveLocation
            }

            $MainObject.FirstReceivedFromIPObject = $FirstReceivedFromIPObject   
        }
        

        #getting data for MapAllIPs Google Maps API Polyline output
        if ($AllReceivedFromIPMap)
        {
            $AllReceivedFromIPObject = Create-AllReceivedFromIPObject -MessageObject $MainObject -SavePath $LogLocation

            $MainObject.AllReceivedFromIPObject = $AllReceivedFromIPObject 
        }

    }
    End
    {
        return $MainObject
    }
}
