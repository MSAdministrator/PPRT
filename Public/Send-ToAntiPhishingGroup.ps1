#requires -Version 2
function Send-ToAntiPhishingGroup
{
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true,
        HelpMessage = 'Please the trimmed phishing url link')]
        [string]$trimmedlink,

        [parameter(Mandatory = $true,
        HelpMessage = 'Please provide th Send On Behalf email address.')]
        $From,

        [parameter(Mandatory = $true,
        HelpMessage = 'Please provide a logging location')]
        $LogLocation
    ) 

    $date = Get-Date -Format yyyyMMdd
    "$trimmedlink" + ',' + "$date"
    
    try
    {
        $outlook = New-Object -ComObject Outlook.Application
        $Mail = $outlook.CreateItem(0)
        $Mail.To = 'anti-phishing-email-reply-discuss@googlegroups.com'
        $Mail.Sentonbehalfofname = "$($From)"
        $Mail.Subject = ('Phishing Links ' + $date)
        $Mail.Body = "$trimmedlink" + ',' + "$date"
        $Mail.Send()
        Write-LogEntry -type Info -message 'Sucessfully sent notification to Abuse Contact' -Folder $LogLocation -CustomMessage "Sent $("$trimmedlink" + ',' + "$date") to: anti-phishing-email-reply-discuss@googlegroups.com"
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
