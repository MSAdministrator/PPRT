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
        $From
    ) 

    $outlook = New-Object -ComObject Outlook.Application
    $Mail = $outlook.CreateItem(0)
    $Mail.To = 'spam@access.ironport.com;phishing-report@us-cert.gov;spam@uce.gov'
    $Mail.Attachments.Add($messagetoattach)
    $Mail.Sentonbehalfofname = "$($From)"
    $Mail.Subject = 'Phishing E-Mail'
    $Mail.Body = "The attached email is a phishing email: $originallink"
    $Mail.Send()
}
