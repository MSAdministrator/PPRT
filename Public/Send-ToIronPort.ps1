#requires -Version 2
function Send-ToIronPort
{
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true,
        HelpMessage = 'Please the original phishing url link')]
        [string]$originallink,

        [parameter(Mandatory = $true,
        HelpMessage = 'Please provide the original message to attach to email. ')]
        $messagetoattach,

        [parameter(Mandatory = $true,
        HelpMessage = 'Please provide th Send On Behalf email address.')]
        $From,

        [parameter(Mandatory = $true,
        HelpMessage = 'Please provide a logging location')]
        $LogLocation
    ) 

    try
    {
        $outlook = New-Object -ComObject Outlook.Application
        $Mail = $outlook.CreateItem(0)
        $Mail.To = 'spam@access.ironport.com;phishing-report@us-cert.gov;spam@uce.gov'
        $Mail.Attachments.Add($messagetoattach)
        $Mail.Sentonbehalfofname = "$($From)"
        $Mail.Subject = 'Phishing E-Mail'
        $Mail.Body = "The attached email is a phishing email: $originallink"
        $Mail.Send()
        Write-LogEntry -type Info -message 'Sucessfully sent notification to Abuse Contact' -Folder $LogLocation -CustomMessage 'Sent to: spam@access.ironport.com;phishing-report@us-cert.gov;spam@uce.gov'
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
