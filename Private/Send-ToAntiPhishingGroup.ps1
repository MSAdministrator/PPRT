#requires -Version 2
function Send-ToAntiPhishingGroup
{
    param (
        [parameter(Mandatory = $true,
        HelpMessage = 'Please the trimmed phishing url link')]
        [string]$trimmedlink,

        [parameter(Mandatory = $true,
        HelpMessage = 'Please provide th Send On Behalf email address.')]
        $From
    ) 

    $date = Get-Date -Format yyyyMMdd
    "$trimmedlink" + ',' + "$date"
       
    $outlook = New-Object -ComObject Outlook.Application
    $Mail = $outlook.CreateItem(0)
    $Mail.To = 'anti-phishing-email-reply-discuss@googlegroups.com'
    $Mail.Sentonbehalfofname = "$($From)"
    $Mail.Subject = ('Phishing Links ' + $date)
    $Mail.Body = "$trimmedlink" + ',' + "$date"
    $Mail.Send()
}
