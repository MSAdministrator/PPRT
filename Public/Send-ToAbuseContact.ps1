#requires -Version 2
function Send-ToAbuseContact
{ 	
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true,
        HelpMessage = 'Please the original phishing url link')]
        [string]$originallink,

        [parameter(Mandatory = $true,
        HelpMessage = 'Please provide the abuse contact to send email to.')]
        $abusecontact,

        [parameter(Mandatory = $true,
        HelpMessage = 'Please provide the original message to attach to email. ')]
        $messagetoattach,

        [parameter(Mandatory = $true,
        HelpMessage = 'Please provide the Send On Behalf email address.')]
        $From,

        [parameter(Mandatory = $true,
        HelpMessage = 'Please provide a logging location')]
        $LogLocation
    ) 
    
    try
    {
        $outlook = New-Object -ComObject Outlook.Application
        $Mail = $outlook.CreateItem(0)
        $Mail.To = "$abusecontact"
        $Mail.Attachments.Add($messagetoattach)
        $Mail.Sentonbehalfofname = "$($From)"
        $Mail.Subject = 'Remove Phishing Website'
        $Mail.Body = "We have received a phishing attempt (attached) that is using an IP registered to this contact.  Please remove this site as soon as you can: $originallink'.' `n`nIn addition, any logs you can provide surrounding the registration or usage of this site would help us understand who is targeting our environment.`n`n Thank you!"
        $Mail.Send()
        Write-LogEntry -type Info -message 'Sucessfully sent notification to Abuse Contact' -Folder $LogLocation -CustomMessage "Sent to: $abusecontact"
        return $true
    }
    catch
    {
        $msg = ('An error occurred that could not be resolved: {0}' -f $_.Exception.Message)
        Write-LogEntry -type ERROR -message 'Error Sending to Abuse Contact' -Folder $LogLocation -CustomMessage "$msg"
        Write-LogEntry -type ERROR -message 'Exception' -Folder $LogLocation -CustomMessage "$($_.Exception)"
        Write-LogEntry -type ERROR -message 'Unknown Exception' -Folder $LogLocation -CustomMessage "$($_)"
        return $false
    }
}
